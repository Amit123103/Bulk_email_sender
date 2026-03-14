"""
SMTPEngine v2.0 — Advanced email sending engine.
Features: Anti-spam headers, smart throttling with jitter, provider-aware rate limiting,
bounce classification, connection keep-alive, batch chunking, warm-up mode,
unsubscribe link injection, plain-text fallback, DKIM-friendly Message-IDs.
"""

import smtplib
import ssl
import re
import time
import random
import logging
import uuid
import os
import html
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate, formataddr, make_msgid
from datetime import datetime
from typing import Optional, Callable
from html.parser import HTMLParser


class HTMLStripper(HTMLParser):
    """Strip HTML tags to produce plain text fallback."""
    def __init__(self):
        super().__init__()
        self.result = []
        self.skip = False
    def handle_starttag(self, tag, attrs):
        if tag in ('style', 'script'):
            self.skip = True
        elif tag == 'br':
            self.result.append('\n')
        elif tag in ('p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'tr', 'li'):
            self.result.append('\n')
    def handle_endtag(self, tag):
        if tag in ('style', 'script'):
            self.skip = False
        elif tag in ('p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
            self.result.append('\n')
    def handle_data(self, data):
        if not self.skip:
            self.result.append(data)
    def get_text(self):
        return re.sub(r'\n{3,}', '\n\n', ''.join(self.result)).strip()


def html_to_plain(html_body: str) -> str:
    """Convert HTML email body to plain text."""
    try:
        stripper = HTMLStripper()
        stripper.feed(html.unescape(html_body))
        return stripper.get_text()
    except Exception:
        return re.sub(r'<[^>]+>', '', html_body)


class SMTPEngine:
    """
    Advanced SMTP engine with anti-spam headers, smart throttling,
    provider-aware rate limiting, bounce classification, batch chunking,
    warm-up mode, and connection keep-alive.
    """

    VARIABLE_PATTERN = re.compile(r'\{\{(\w+)\}\}')
    MAX_RETRIES = 2

    # Provider-specific rate limits (emails per minute)
    PROVIDER_RATES = {
        'smtp.gmail.com': 20,
        'smtp-mail.outlook.com': 15,
        'smtp.mail.yahoo.com': 10,
        'smtp.zoho.com': 15,
        'smtp.sendgrid.net': 50,
        'smtp.office365.com': 15,
        'smtp.aol.com': 10,
    }

    # Hard bounce SMTP error codes (permanent failure — don't retry)
    HARD_BOUNCE_CODES = {550, 551, 552, 553, 554}
    # Soft bounce codes (temporary — retry later)
    SOFT_BOUNCE_CODES = {421, 450, 451, 452}

    # Batch settings
    DEFAULT_BATCH_SIZE = 50
    BATCH_COOLDOWN = 10  # seconds between batches

    def __init__(self):
        self.connections = {}     # account_id → smtplib connection
        self.last_send_time = {}  # account_id → timestamp of last send
        self.send_counts = {}     # account_id → list of timestamps (for rate tracking)
        self.bounce_stats = {}    # account_id → {"hard": 0, "soft": 0, "sent": 0}
        self.logger = logging.getLogger('SMTPEngine')
        self.logger.setLevel(logging.DEBUG)
        self._warm_up = False
        self._warm_up_counter = 0

    # ─── Connection Management ────────────────────────────────────

    def build_connection(self, account: dict) -> smtplib.SMTP:
        """Create and authenticate an SMTP connection for an account."""
        host = account["smtp_host"].strip()
        port = int(account["smtp_port"])
        security = account.get("smtp_security", "TLS")
        email = account["email"]
        password = account["password"]

        # User might have copied App Password with spaces or newlines
        password_clean = "".join(password.split())

        try:
            if security == "SSL":
                ctx = ssl.create_default_context()
                server = smtplib.SMTP_SSL(host, port, context=ctx, timeout=30)
            else:
                server = smtplib.SMTP(host, port, timeout=30)
                server.ehlo()
                if security == "TLS":
                    ctx = ssl.create_default_context()
                    server.starttls(context=ctx)
                    server.ehlo()

            server.login(email, password_clean)
            self.logger.info(f"Connected to {host}:{port} as {email}")
            return server

        except smtplib.SMTPAuthenticationError as e:
            error_msg = e.smtp_error.decode('utf-8', errors='replace') if isinstance(e.smtp_error, bytes) else str(e.smtp_error)
            raise ConnectionError(
                f"Authentication failed for {email}: {e.smtp_code} {error_msg}"
            )
        except smtplib.SMTPConnectError as e:
            raise ConnectionError(f"Cannot connect to {host}:{port}: {e}")
        except smtplib.SMTPException as e:
            raise ConnectionError(f"SMTP error for {email}: {e}")
        except Exception as e:
            raise ConnectionError(f"Connection failed for {email}: {e}")

    def connect_account(self, account_id: str, account: dict) -> bool:
        """Open and store an SMTP connection for the given account."""
        try:
            conn = self.build_connection(account)
            self.connections[account_id] = conn
            if account_id not in self.bounce_stats:
                self.bounce_stats[account_id] = {"hard": 0, "soft": 0, "sent": 0}
            return True
        except ConnectionError as e:
            self.logger.error(f"Connect failed [{account_id}]: {e}")
            return False

    def disconnect_account(self, account_id: str):
        """Close and remove a stored connection."""
        conn = self.connections.pop(account_id, None)
        if conn:
            try:
                conn.quit()
            except Exception:
                try:
                    conn.close()
                except Exception:
                    pass

    def disconnect_all(self):
        """Close all active connections."""
        for aid in list(self.connections.keys()):
            self.disconnect_account(aid)

    def is_connection_alive(self, account_id: str) -> bool:
        """Check if a connection is still alive via NOOP."""
        conn = self.connections.get(account_id)
        if not conn:
            return False
        try:
            status = conn.noop()
            return status[0] == 250
        except Exception:
            return False

    def keep_alive(self, account_id: str):
        """Send NOOP to keep connection alive during pauses."""
        conn = self.connections.get(account_id)
        if conn:
            try:
                conn.noop()
            except Exception:
                pass

    def ensure_connection(self, account_id: str, account: dict) -> bool:
        """Reconnect if dead, up to 3 attempts with exponential backoff."""
        if self.is_connection_alive(account_id):
            return True

        self.disconnect_account(account_id)
        for attempt in range(3):
            try:
                conn = self.build_connection(account)
                self.connections[account_id] = conn
                self.logger.info(f"Reconnected {account['email']} (attempt {attempt + 1})")
                return True
            except ConnectionError:
                wait = (attempt + 1) * 3  # 3s, 6s, 9s
                if attempt < 2:
                    time.sleep(wait)
        return False

    # ─── Rate Limiting & Throttling ───────────────────────────────

    def get_provider_rate(self, smtp_host: str) -> int:
        """Get max emails per minute for a provider."""
        return self.PROVIDER_RATES.get(smtp_host.lower(), 30)

    def calculate_delay(self, base_delay: float, smtp_host: str) -> float:
        """Calculate delay with ±30% jitter and provider rate limit enforcement."""
        # Provider minimum delay
        rate = self.get_provider_rate(smtp_host)
        provider_min_delay = 60.0 / rate  # seconds per email

        # Apply jitter (±30%)
        jitter = base_delay * 0.3
        actual_delay = base_delay + random.uniform(-jitter, jitter)

        # Enforce provider minimum
        actual_delay = max(actual_delay, provider_min_delay)

        # Warm-up mode: start slow
        if self._warm_up and self._warm_up_counter < 20:
            # Gradually decrease from 60s to base_delay over first 20 emails
            warm_factor = max(0.1, 1.0 - (self._warm_up_counter / 20.0))
            warm_delay = 60.0 * warm_factor
            actual_delay = max(actual_delay, warm_delay)

        return round(actual_delay, 2)

    def enforce_rate_limit(self, account_id: str, smtp_host: str):
        """Wait if sending too fast for the provider rate limit."""
        now = time.time()
        rate = self.get_provider_rate(smtp_host)
        window = 60.0  # 1-minute window

        if account_id not in self.send_counts:
            self.send_counts[account_id] = []

        # Clean old timestamps
        self.send_counts[account_id] = [
            t for t in self.send_counts[account_id] if now - t < window
        ]

        if len(self.send_counts[account_id]) >= rate:
            # Wait for oldest timestamp to age out
            oldest = self.send_counts[account_id][0]
            wait = window - (now - oldest) + 0.5
            if wait > 0:
                self.logger.info(f"Rate limit: waiting {wait:.1f}s for {smtp_host}")
                time.sleep(wait)

        self.send_counts[account_id].append(time.time())

    # ─── Personalization ──────────────────────────────────────────

    CONDITIONAL_PATTERN = re.compile(
        r'\{\{#if\s+(\w+)\}\}(.*?)(?:\{\{#else\}\}(.*?))?\{\{/if\}\}',
        re.DOTALL
    )

    def personalize(self, template: str, recipient: dict) -> str:
        """Replace {{variable}} placeholders and process {{#if}} conditionals."""
        # 1. Process conditional blocks first
        def conditional_replacer(match):
            var_name = match.group(1).lower()
            if_content = match.group(2)
            else_content = match.group(3) or ""

            # Check if variable exists and is truthy
            value = None
            for key, v in recipient.items():
                if key.lower() == var_name or key.lower().replace(' ', '_') == var_name:
                    value = v
                    break

            if value and str(value).strip():
                return if_content
            return else_content

        result = self.CONDITIONAL_PATTERN.sub(conditional_replacer, template)

        # 2. Replace {{variable}} placeholders
        def replacer(match):
            var = match.group(1).lower()

            # Built-in special variables
            if var == "first_name":
                name = recipient.get("name", "") or recipient.get("full_name", "")
                return name.split()[0] if name and name.strip() else ""
            elif var == "last_name":
                name = recipient.get("name", "") or recipient.get("full_name", "")
                parts = name.split()
                return parts[-1] if len(parts) > 1 else ""
            elif var == "date":
                return datetime.now().strftime("%B %d, %Y")
            elif var == "time":
                return datetime.now().strftime("%I:%M %p")
            elif var == "year":
                return str(datetime.now().year)
            elif var == "unsubscribe_link":
                return recipient.get("_unsubscribe_url", "#unsubscribe")

            # Look up in recipient dict (case-insensitive)
            for key, value in recipient.items():
                if key.lower() == var or key.lower().replace(' ', '_') == var:
                    return str(value) if value else ""
            return ""

        return self.VARIABLE_PATTERN.sub(replacer, result)

    def inject_tracking_pixel(self, html_body: str, tracking_id: str,
                               tracking_url: str = None) -> str:
        """Inject a 1x1 tracking pixel into HTML email for open tracking."""
        if not tracking_url:
            tracking_url = f"https://track.example.com/open/{tracking_id}.png"
        pixel = f'<img src="{tracking_url}" width="1" height="1" alt="" style="display:none;border:0;" />'
        # Insert before </body> if exists, else append
        if '</body>' in html_body.lower():
            idx = html_body.lower().rfind('</body>')
            return html_body[:idx] + pixel + html_body[idx:]
        return html_body + pixel

    # ─── Bounce Classification ────────────────────────────────────

    def classify_bounce(self, error: Exception) -> str:
        """Classify SMTP error as hard bounce, soft bounce, or transient."""
        error_str = str(error)
        # Extract SMTP code
        code = None
        code_match = re.search(r'\b(\d{3})\b', error_str)
        if code_match:
            code = int(code_match.group(1))

        if code in self.HARD_BOUNCE_CODES:
            return "hard_bounce"
        elif code in self.SOFT_BOUNCE_CODES:
            return "soft_bounce"

        # Pattern matching
        hard_patterns = ['does not exist', 'user unknown', 'mailbox not found',
                         'no such user', 'invalid recipient', 'address rejected',
                         'recipient rejected', 'mailbox unavailable']
        soft_patterns = ['mailbox full', 'over quota', 'temporarily', 'try again',
                         'too many connections', 'rate limit', 'service unavailable']

        error_lower = error_str.lower()
        for p in hard_patterns:
            if p in error_lower:
                return "hard_bounce"
        for p in soft_patterns:
            if p in error_lower:
                return "soft_bounce"

        return "transient"

    def update_bounce_stats(self, account_id: str, bounce_type: str):
        """Track bounce stats per account."""
        if account_id not in self.bounce_stats:
            self.bounce_stats[account_id] = {"hard": 0, "soft": 0, "sent": 0}
        if bounce_type == "hard_bounce":
            self.bounce_stats[account_id]["hard"] += 1
        elif bounce_type == "soft_bounce":
            self.bounce_stats[account_id]["soft"] += 1

    def get_account_health(self, account_id: str) -> float:
        """Return health score 0-100 for an account based on bounce rate."""
        stats = self.bounce_stats.get(account_id, {"hard": 0, "soft": 0, "sent": 0})
        total = stats["sent"] + stats["hard"] + stats["soft"]
        if total < 5:
            return 100.0  # Not enough data
        bounces = stats["hard"] * 2 + stats["soft"]  # Hard bounces count double
        score = max(0, 100 - (bounces / total * 100))
        return round(score, 1)

    # ─── MIME Message Builder ─────────────────────────────────────

    def build_mime_message(self, sender_email: str, sender_name: str,
                           recipient_email: str, subject: str, body: str,
                           is_html: bool = True,
                           attachment_paths: list = None,
                           cc: str = '', bcc: str = '',
                           reply_to: str = None,
                           unsubscribe_url: str = None) -> MIMEMultipart:
        """Build a complete MIME message with anti-spam headers, plain-text fallback, and attachments."""
        sender_domain = sender_email.split('@')[1] if '@' in sender_email else 'localhost'

        # Ensure attachment_paths is a list (for backward compatibility if someone passes a string)
        if isinstance(attachment_paths, str):
            attachment_paths = [attachment_paths]

        if attachment_paths and any(os.path.isfile(p) for p in attachment_paths):
            msg = MIMEMultipart('mixed')
            body_container = MIMEMultipart('alternative')
            msg.attach(body_container)
        else:
            msg = MIMEMultipart('alternative')
            body_container = msg

        # ── Essential Headers ──
        if sender_name:
            msg['From'] = formataddr((sender_name, sender_email))
        else:
            msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg['Date'] = formatdate(localtime=True)

        # RFC-compliant Message-ID
        msg['Message-ID'] = make_msgid(domain=sender_domain)

        # ── Anti-Spam Headers ──
        msg['Reply-To'] = reply_to or sender_email
        msg['Return-Path'] = sender_email
        msg['MIME-Version'] = '1.0'
        msg['X-Mailer'] = 'BulkEmailPro/3.0'
        msg['X-Priority'] = '3'  # Normal priority
        msg['Precedence'] = 'bulk'

        # Unsubscribe header (improves deliverability significantly)
        if unsubscribe_url:
            msg['List-Unsubscribe'] = f'<{unsubscribe_url}>'
            msg['List-Unsubscribe-Post'] = 'List-Unsubscribe=One-Click'

        if cc:
            msg['Cc'] = cc
        # Note: BCC is NOT added as a header (it's only used in envelope)

        # ── Always attach PLAIN TEXT first, then HTML ──
        if is_html:
            plain_text = html_to_plain(body)
            body_container.attach(MIMEText(plain_text, 'plain', 'utf-8'))
            body_container.attach(MIMEText(body, 'html', 'utf-8'))
        else:
            body_container.attach(MIMEText(body, 'plain', 'utf-8'))

        # ── Attachments ──
        if attachment_paths:
            for attachment_path in attachment_paths:
                if attachment_path and os.path.isfile(attachment_path):
                    try:
                        file_size = os.path.getsize(attachment_path)
                        if file_size > 25 * 1024 * 1024:  # 25MB limit
                            self.logger.warning(f"Attachment too large: {file_size/1024/1024:.1f}MB (max 25MB)")
                        else:
                            with open(attachment_path, 'rb') as f:
                                part = MIMEBase('application', 'octet-stream')
                                part.set_payload(f.read())
                                encoders.encode_base64(part)
                                filename = os.path.basename(attachment_path)
                                part.add_header(
                                    'Content-Disposition',
                                    f'attachment; filename="{filename}"'
                                )
                                msg.attach(part)
                    except IOError as e:
                        self.logger.warning(f"Could not attach file {attachment_path}: {e}")

        return msg

    # ─── Send Single Email ────────────────────────────────────────

    def send_one(self, account_id: str, account: dict, recipient: dict,
                 subject_template: str, body_template: str,
                 is_html: bool = True, attachment_paths: list = None,
                 cc: str = '', bcc: str = '',
                 reply_to: str = None,
                 unsubscribe_url: str = None) -> dict:
        """Send a single email with personalization and anti-spam headers. Return result dict."""
        timestamp = datetime.now().isoformat()
        email_to = recipient.get("_email", "")

        try:
            # Personalize
            subject = self.personalize(subject_template, recipient)
            body = self.personalize(body_template, recipient)

            # Ensure connection
            if not self.ensure_connection(account_id, account):
                return {
                    "_email": email_to,
                    "account_used": account["email"],
                    "account_id": account_id,
                    "status": "failed",
                    "bounce_type": "transient",
                    "timestamp": timestamp,
                    "error": "Could not establish SMTP connection"
                }

            # Enforce rate limit before sending
            self.enforce_rate_limit(account_id, account.get("smtp_host", ""))

            # Build message with anti-spam headers
            sender_name = account.get("nickname", "")
            msg = self.build_mime_message(
                sender_email=account["email"],
                sender_name=sender_name,
                recipient_email=email_to,
                subject=subject,
                body=body,
                is_html=is_html,
                attachment_paths=attachment_paths,
                cc=cc,
                bcc=bcc,
                reply_to=reply_to or account.get("reply_to", account["email"]),
                unsubscribe_url=unsubscribe_url
            )

            # Build recipient list (envelope)
            recipients = [email_to]
            if cc:
                recipients.extend([e.strip() for e in cc.split(',') if e.strip()])
            if bcc:
                recipients.extend([e.strip() for e in bcc.split(',') if e.strip()])

            self.connections[account_id].sendmail(
                account["email"], recipients, msg.as_string()
            )

            # Update bounce stats
            if account_id in self.bounce_stats:
                self.bounce_stats[account_id]["sent"] += 1

            self.last_send_time[account_id] = time.time()

            return {
                "_email": email_to,
                "account_used": account["email"],
                "account_id": account_id,
                "status": "sent",
                "bounce_type": None,
                "timestamp": datetime.now().isoformat(),
                "error": None
            }

        except smtplib.SMTPRecipientsRefused as e:
            bounce_type = self.classify_bounce(e)
            self.update_bounce_stats(account_id, bounce_type)
            return {
                "_email": email_to,
                "account_used": account["email"],
                "account_id": account_id,
                "status": "failed",
                "bounce_type": bounce_type,
                "timestamp": datetime.now().isoformat(),
                "error": f"Recipient refused: {e}"
            }
        except smtplib.SMTPDataError as e:
            bounce_type = self.classify_bounce(e)
            self.update_bounce_stats(account_id, bounce_type)
            return {
                "_email": email_to,
                "account_used": account["email"],
                "account_id": account_id,
                "status": "failed",
                "bounce_type": bounce_type,
                "timestamp": datetime.now().isoformat(),
                "error": f"SMTP data error: {e}"
            }
        except Exception as e:
            bounce_type = self.classify_bounce(e)
            self.update_bounce_stats(account_id, bounce_type)
            return {
                "_email": email_to,
                "account_used": account["email"],
                "account_id": account_id,
                "status": "failed",
                "bounce_type": bounce_type,
                "timestamp": datetime.now().isoformat(),
                "error": str(e)
            }

    # ─── Bulk Send with Rotation + Batching ───────────────────────

    def send_bulk(self,
                  email_list: list,
                  account_manager,
                  subject_template: str,
                  body_template: str,
                  is_html: bool = True,
                  delay: float = 1.0,
                  use_rotation: bool = True,
                  attachment_paths: list = None,
                  cc: str = '', bcc: str = '',
                  reply_to: str = None,
                  unsubscribe_url: str = None,
                  batch_size: int = None,
                  warm_up: bool = False,
                  progress_callback: Callable = None,
                  stop_flag: Callable = None,
                  pause_flag: Callable = None) -> list:
        """
        Master bulk send with:
        - Account rotation (cycle through available accounts on limit)
        - Smart throttling (jitter + provider-aware rate limits)
        - Batch chunking with cool-down periods
        - Warm-up mode for new accounts
        - Bounce classification
        - Connection keep-alive during pauses
        - Retry with escalating backoff
        """
        results = []
        total = len(email_list)

        if total == 0:
            return results

        # Warm-up mode
        self._warm_up = warm_up
        self._warm_up_counter = 0

        # Batch size
        effective_batch = batch_size or self.DEFAULT_BATCH_SIZE

        # Get rotation order
        rotation_ids = account_manager.get_rotation_order()
        if not rotation_ids:
            if progress_callback:
                progress_callback(0, total, {
                    "_email": "", "status": "failed",
                    "error": "No available email accounts", "account_used": ""
                })
            return results

        current_idx = 0
        current_account_id = rotation_ids[current_idx]
        current_account = account_manager.get_account(current_account_id)

        # Connect first account
        if not self.connect_account(current_account_id, current_account):
            account_manager.update_account_status(
                current_account_id, "failed", "Initial connection failed"
            )
            # Try next account
            switched = False
            while current_idx < len(rotation_ids) - 1:
                current_idx += 1
                current_account_id = rotation_ids[current_idx]
                current_account = account_manager.get_account(current_account_id)
                if self.connect_account(current_account_id, current_account):
                    switched = True
                    break
                account_manager.update_account_status(
                    current_account_id, "failed", "Connection failed"
                )
            if not switched:
                return results

        for i, recipient in enumerate(email_list):
            # ── Check stop ──
            if stop_flag and stop_flag():
                self.logger.info(f"Campaign stopped at email {i + 1}/{total}")
                break

            # ── Pause loop with keep-alive ──
            pause_counter = 0
            while pause_flag and pause_flag():
                time.sleep(0.3)
                pause_counter += 1
                # Send NOOP every ~60 seconds to keep connection alive
                if pause_counter % 200 == 0:
                    self.keep_alive(current_account_id)
                if stop_flag and stop_flag():
                    break

            if stop_flag and stop_flag():
                break

            # ── Batch cool-down ──
            if i > 0 and i % effective_batch == 0:
                cooldown = self.BATCH_COOLDOWN + random.uniform(0, 5)
                self.logger.info(f"Batch cool-down: {cooldown:.1f}s after {i} emails")
                if progress_callback:
                    progress_callback(i, total, {
                        "_email": "", "status": "info",
                        "account_used": current_account["email"],
                        "error": None,
                        "message": f"Batch cool-down ({cooldown:.0f}s) after {i} emails..."
                    })
                time.sleep(cooldown)

            # ── Check if current account still has capacity ──
            acc_data = account_manager.accounts.get(current_account_id, {})
            if acc_data.get("sent_today", 0) >= acc_data.get("daily_limit", 500):
                rotated = False
                if use_rotation:
                    rotation_ids = account_manager.get_rotation_order()
                    for rid in rotation_ids:
                        if rid != current_account_id:
                            next_acc = account_manager.get_account(rid)
                            if next_acc and self.connect_account(rid, next_acc):
                                old_email = current_account["email"]
                                current_account_id = rid
                                current_account = next_acc
                                rotated = True
                                self.logger.info(
                                    f"Rotated from {old_email} to {next_acc['email']}"
                                )
                                if progress_callback:
                                    progress_callback(i, total, {
                                        "_email": "",
                                        "status": "rotation",
                                        "account_used": next_acc["email"],
                                        "error": None,
                                        "message": f"Switched to {next_acc['email']} (previous at daily limit)"
                                    })
                                break

                if not rotated:
                    self.logger.warning("All accounts at daily limit.")
                    if progress_callback:
                        progress_callback(i, total, {
                            "_email": "", "status": "warning",
                            "account_used": "",
                            "error": "All accounts at daily limit. Waiting for reset...",
                        })
                    # Wait loop (max 24 hours)
                    for _ in range(1440):
                        time.sleep(60)
                        account_manager.reset_daily_counters()
                        rotation_ids = account_manager.get_rotation_order()
                        if rotation_ids:
                            current_account_id = rotation_ids[0]
                            current_account = account_manager.get_account(current_account_id)
                            if self.connect_account(current_account_id, current_account):
                                break
                        if stop_flag and stop_flag():
                            break
                    else:
                        break

            # ── Check account health — auto-disable if too many bounces ──
            health = self.get_account_health(current_account_id)
            if health < 30 and self.bounce_stats.get(current_account_id, {}).get("sent", 0) >= 10:
                self.logger.warning(f"Account health low ({health}%), rotating away")
                account_manager.update_account_status(
                    current_account_id, "failed", f"Auto-disabled: health score {health}%"
                )
                # Try to rotate
                if use_rotation:
                    rotation_ids = account_manager.get_rotation_order()
                    switched = False
                    for rid in rotation_ids:
                        if rid != current_account_id:
                            next_acc = account_manager.get_account(rid)
                            if next_acc and self.connect_account(rid, next_acc):
                                current_account_id = rid
                                current_account = next_acc
                                switched = True
                                break
                    if not switched:
                        break

            # ── Send with retries ──
            result = None
            retries = 0
            while retries <= self.MAX_RETRIES:
                result = self.send_one(
                    current_account_id, current_account, recipient,
                    subject_template, body_template,
                    is_html, attachment_paths, cc, bcc,
                    reply_to, unsubscribe_url
                )

                if result["status"] == "sent":
                    account_manager.increment_sent_count(current_account_id)
                    if self._warm_up:
                        self._warm_up_counter += 1
                    break

                bounce_type = result.get("bounce_type", "transient")
                error_str = str(result.get("error", "")).lower()

                # Hard bounce — don't retry, skip this recipient
                if bounce_type == "hard_bounce":
                    self.logger.info(f"Hard bounce for {recipient.get('_email')}, skipping")
                    break

                # Recipient refused — don't retry
                if "recipient" in error_str and "refused" in error_str:
                    break

                # Auth error — rotate to next account
                if "authentication" in error_str or "auth" in error_str:
                    account_manager.update_account_status(
                        current_account_id, "failed", result.get("error")
                    )
                    if use_rotation:
                        rotation_ids = account_manager.get_rotation_order()
                        switched = False
                        for rid in rotation_ids:
                            if rid != current_account_id:
                                next_acc = account_manager.get_account(rid)
                                if next_acc and self.connect_account(rid, next_acc):
                                    current_account_id = rid
                                    current_account = next_acc
                                    switched = True
                                    break
                        if not switched:
                            break
                    retries += 1
                    continue

                # Disconnect or connection reset — reconnect with backoff
                if ("disconnect" in error_str or "connection" in error_str
                        or "reset" in error_str or "broken" in error_str):
                    self.disconnect_account(current_account_id)
                    backoff = (retries + 1) * 3
                    time.sleep(backoff)
                    retries += 1
                    continue

                # Rate limit error — wait and retry
                if "rate" in error_str or "too many" in error_str or "throttl" in error_str:
                    wait_time = 30 * (retries + 1)
                    self.logger.info(f"Rate limited, waiting {wait_time}s")
                    time.sleep(wait_time)
                    retries += 1
                    continue

                # Other error — retry with backoff
                retries += 1
                if retries <= self.MAX_RETRIES:
                    time.sleep(2 * retries)

            if result:
                results.append(result)
                if progress_callback:
                    progress_callback(i + 1, total, result)

            # Smart delay between emails
            if i < total - 1 and delay > 0:
                actual_delay = self.calculate_delay(delay, current_account.get("smtp_host", ""))
                time.sleep(actual_delay)

        # Cleanup
        self.disconnect_all()
        self._warm_up = False
        return results

    # ─── Campaign Summary ─────────────────────────────────────────

    def get_campaign_summary(self, results: list) -> dict:
        """Generate detailed summary from campaign results."""
        sent = [r for r in results if r.get("status") == "sent"]
        failed = [r for r in results if r.get("status") == "failed"]
        hard = [r for r in failed if r.get("bounce_type") == "hard_bounce"]
        soft = [r for r in failed if r.get("bounce_type") == "soft_bounce"]
        transient = [r for r in failed if r.get("bounce_type") == "transient"]

        # Account usage breakdown
        account_usage = {}
        for r in results:
            acc = r.get("account_used", "unknown")
            if acc not in account_usage:
                account_usage[acc] = {"sent": 0, "failed": 0}
            if r.get("status") == "sent":
                account_usage[acc]["sent"] += 1
            else:
                account_usage[acc]["failed"] += 1

        # Domain distribution
        domain_dist = {}
        for r in sent:
            email = r.get("_email", "")
            if "@" in email:
                domain = email.split("@")[1]
                domain_dist[domain] = domain_dist.get(domain, 0) + 1

        total = len(results)
        return {
            "total": total,
            "sent": len(sent),
            "failed": len(failed),
            "hard_bounces": len(hard),
            "soft_bounces": len(soft),
            "transient_errors": len(transient),
            "success_rate": round(len(sent) / max(total, 1) * 100, 1),
            "account_usage": account_usage,
            "domain_distribution": dict(sorted(domain_dist.items(), key=lambda x: -x[1])[:10]),
            "bounce_stats": dict(self.bounce_stats),
        }
