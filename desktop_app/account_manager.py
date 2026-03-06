"""
AccountManager v2.0 — Enhanced email account management.
Features: Fernet AES-128 encryption, CRUD, daily limits, smart rotation,
health scoring, cooldown tracking, auto-disable, import/export backup.
"""

import json
import os
import uuid
import smtplib
import ssl
import socket
import threading
import time
import base64
from datetime import datetime, date
from pathlib import Path
from typing import Optional

try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
except ImportError:
    CRYPTO_AVAILABLE = False

try:
    import keyring
    KEYRING_AVAILABLE = True
except ImportError:
    KEYRING_AVAILABLE = False


class AccountManager:
    """
    Manages all email accounts with encryption, health tracking, and smart rotation.
    Each account tracks: id, nickname, email, SMTP config, encrypted password,
    daily_limit, sent_today, health_score, cooldown, status, errors.
    """

    ACCOUNTS_FILE = "accounts.json"
    KEY_FILE = "secret.key"

    # Cooldown: minimum seconds between sends per account
    MIN_SEND_GAP = 2.0

    def __init__(self, data_dir: str = None):
        self.data_dir = Path(data_dir) if data_dir else Path(".")
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.accounts_path = self.data_dir / self.ACCOUNTS_FILE
        self.key_path = self.data_dir / self.KEY_FILE
        self.accounts = {}
        self._lock = threading.Lock()

        if CRYPTO_AVAILABLE:
            self.key = self.load_key()
            self.cipher = Fernet(self.key)
        else:
            self.key = None
            self.cipher = None

        self.load_accounts()
        self.reset_daily_counters()

    # ─── Encryption ────────────────────────────────────────────────

    def generate_key(self) -> bytes:
        key = Fernet.generate_key()
        try:
            with open(self.key_path, "wb") as f:
                f.write(key)
        except IOError as e:
            raise RuntimeError(f"Failed to save encryption key: {e}")
        return key

    def load_key(self) -> bytes:
        if self.key_path.exists():
            try:
                with open(self.key_path, "rb") as f:
                    return f.read()
            except IOError:
                return self.generate_key()
        return self.generate_key()

    def encrypt_password(self, password: str) -> str:
        if not CRYPTO_AVAILABLE or not self.cipher:
            return "b64:" + base64.b64encode(password.encode("utf-8")).decode("utf-8")
        encrypted = self.cipher.encrypt(password.encode("utf-8"))
        return encrypted.decode("utf-8")

    def decrypt_password(self, encrypted: str) -> str:
        if not encrypted:
            return ""
        if encrypted.startswith("b64:"):
            return base64.b64decode(encrypted[4:].encode("utf-8")).decode("utf-8")
        if not CRYPTO_AVAILABLE or not self.cipher:
            return encrypted
        try:
            decrypted = self.cipher.decrypt(encrypted.encode("utf-8"))
            return decrypted.decode("utf-8")
        except Exception:
            return ""

    # ─── CRUD ──────────────────────────────────────────────────────

    def add_account(self, nickname: str, email: str, smtp_host: str,
                    smtp_port: int, smtp_security: str, password: str,
                    daily_limit: int = 500) -> dict:
        if not nickname or not nickname.strip():
            raise ValueError("Account nickname is required.")
        if not email or not email.strip():
            raise ValueError("Email address is required.")
        if not smtp_host or not smtp_host.strip():
            raise ValueError("SMTP host is required.")
        if smtp_port < 1 or smtp_port > 65535:
            raise ValueError("SMTP port must be between 1 and 65535.")
        if smtp_security not in ("TLS", "SSL", "None"):
            raise ValueError("SMTP security must be TLS, SSL, or None.")
        if not password:
            raise ValueError("Password is required.")

        # Check duplicate
        for acc in self.accounts.values():
            if acc["email"].lower() == email.strip().lower():
                raise ValueError(f"An account with email '{email}' already exists.")

        account_id = str(uuid.uuid4())[:8]
        account = {
            "id": account_id,
            "nickname": nickname.strip(),
            "email": email.strip(),
            "smtp_host": smtp_host.strip(),
            "smtp_port": int(smtp_port),
            "smtp_security": smtp_security,
            "password": self.encrypt_password(password),
            "daily_limit": max(1, int(daily_limit)),
            "sent_today": 0,
            "last_reset_date": str(date.today()),
            "status": "unknown",
            "last_error": None,
            "is_default": len(self.accounts) == 0,
            "created_at": datetime.now().isoformat(),
            # v2.0 enhancements
            "health_score": 100.0,
            "total_sent": 0,
            "total_failed": 0,
            "hard_bounces": 0,
            "soft_bounces": 0,
            "last_send_time": None,
            "auto_disabled": False,
            "reply_to": email.strip(),
        }

        with self._lock:
            self.accounts[account_id] = account
            self.save_accounts()

        safe = dict(account)
        safe["password"] = "***"
        return safe

    def update_account(self, account_id: str, updates: dict) -> bool:
        with self._lock:
            if account_id not in self.accounts:
                return False
            for key, value in updates.items():
                if key == "password" and value:
                    self.accounts[account_id]["password"] = self.encrypt_password(value)
                elif key in self.accounts[account_id] and key not in ("id", "created_at"):
                    self.accounts[account_id][key] = value
            self.save_accounts()
        return True

    def delete_account(self, account_id: str) -> bool:
        with self._lock:
            if account_id not in self.accounts:
                return False
            was_default = self.accounts[account_id].get("is_default", False)
            del self.accounts[account_id]
            if was_default and self.accounts:
                first_id = next(iter(self.accounts))
                self.accounts[first_id]["is_default"] = True
            self.save_accounts()
        return True

    def get_account(self, account_id: str) -> dict:
        if account_id not in self.accounts:
            return None
        acc = dict(self.accounts[account_id])
        acc["password"] = self.decrypt_password(acc["password"])
        return acc

    def get_account_by_nickname(self, nickname: str) -> dict:
        for acc in self.accounts.values():
            if acc["nickname"].lower() == nickname.lower():
                result = dict(acc)
                result["password"] = self.decrypt_password(result["password"])
                return result
        return None

    def get_all_accounts(self) -> list:
        result = []
        for acc in self.accounts.values():
            safe = dict(acc)
            safe["password"] = "***"
            result.append(safe)
        return result

    def get_available_accounts(self) -> list:
        self.reset_daily_counters()
        available = []
        for acc in self.accounts.values():
            if (acc["sent_today"] < acc["daily_limit"]
                    and acc["status"] != "failed"
                    and not acc.get("auto_disabled", False)):
                safe = dict(acc)
                safe["password"] = "***"
                safe["remaining"] = acc["daily_limit"] - acc["sent_today"]
                available.append(safe)
        available.sort(key=lambda a: a["remaining"], reverse=True)
        return available

    def get_default_account(self) -> dict:
        for acc in self.accounts.values():
            if acc.get("is_default"):
                result = dict(acc)
                result["password"] = self.decrypt_password(result["password"])
                return result
        if self.accounts:
            first = dict(next(iter(self.accounts.values())))
            first["password"] = self.decrypt_password(first["password"])
            return first
        return None

    def set_default_account(self, account_id: str) -> bool:
        with self._lock:
            if account_id not in self.accounts:
                return False
            for aid in self.accounts:
                self.accounts[aid]["is_default"] = (aid == account_id)
            self.save_accounts()
        return True

    # ─── Daily Limit Tracking ─────────────────────────────────────

    def increment_sent_count(self, account_id: str) -> bool:
        with self._lock:
            if account_id not in self.accounts:
                return False
            acc = self.accounts[account_id]
            acc["sent_today"] += 1
            acc["total_sent"] = acc.get("total_sent", 0) + 1
            acc["last_send_time"] = datetime.now().isoformat()
            if acc["sent_today"] >= acc["daily_limit"]:
                acc["status"] = "limit_reached"
            self.save_accounts()
            return acc["sent_today"] < acc["daily_limit"]

    def increment_fail_count(self, account_id: str, bounce_type: str = "transient"):
        """Track failures and bounces per account."""
        with self._lock:
            if account_id not in self.accounts:
                return
            acc = self.accounts[account_id]
            acc["total_failed"] = acc.get("total_failed", 0) + 1
            if bounce_type == "hard_bounce":
                acc["hard_bounces"] = acc.get("hard_bounces", 0) + 1
            elif bounce_type == "soft_bounce":
                acc["soft_bounces"] = acc.get("soft_bounces", 0) + 1
            # Recalculate health score
            self._recalculate_health(account_id)
            self.save_accounts()

    def reset_daily_counters(self):
        today = str(date.today())
        changed = False
        with self._lock:
            for acc in self.accounts.values():
                if acc.get("last_reset_date") != today:
                    acc["sent_today"] = 0
                    acc["last_reset_date"] = today
                    if acc["status"] == "limit_reached":
                        acc["status"] = "unknown"
                    changed = True
            if changed:
                self.save_accounts()

    # ─── Health Scoring ───────────────────────────────────────────

    def _recalculate_health(self, account_id: str):
        """Recalculate health score 0-100 based on success/bounce rate."""
        acc = self.accounts.get(account_id)
        if not acc:
            return
        total_sent = acc.get("total_sent", 0)
        total_failed = acc.get("total_failed", 0)
        hard = acc.get("hard_bounces", 0)
        total = total_sent + total_failed
        if total < 5:
            acc["health_score"] = 100.0
            return
        # Hard bounces are weighted 2x
        penalty = (hard * 2 + total_failed) / total * 100
        score = max(0, 100 - penalty)
        acc["health_score"] = round(score, 1)

        # Auto-disable if health drops below 30% after 10+ sends
        if score < 30 and total >= 10:
            acc["auto_disabled"] = True
            acc["status"] = "failed"
            acc["last_error"] = f"Auto-disabled: health score {score}%"

    def get_account_health(self, account_id: str) -> float:
        """Get health score for an account."""
        acc = self.accounts.get(account_id)
        return acc.get("health_score", 100.0) if acc else 0.0

    def re_enable_account(self, account_id: str) -> bool:
        """Re-enable an auto-disabled account."""
        with self._lock:
            if account_id not in self.accounts:
                return False
            acc = self.accounts[account_id]
            acc["auto_disabled"] = False
            acc["status"] = "unknown"
            acc["last_error"] = None
            acc["health_score"] = 50.0  # Start at 50% after re-enable
            self.save_accounts()
        return True

    # ─── Cooldown Tracking ────────────────────────────────────────

    def get_time_since_last_send(self, account_id: str) -> float:
        """Get seconds since last send for this account."""
        acc = self.accounts.get(account_id)
        if not acc or not acc.get("last_send_time"):
            return float('inf')
        try:
            last = datetime.fromisoformat(acc["last_send_time"])
            return (datetime.now() - last).total_seconds()
        except Exception:
            return float('inf')

    def is_cooldown_ready(self, account_id: str) -> bool:
        """Check if account has waited enough since last send."""
        return self.get_time_since_last_send(account_id) >= self.MIN_SEND_GAP

    # ─── Account Status ───────────────────────────────────────────

    def update_account_status(self, account_id: str, status: str, error: str = None):
        with self._lock:
            if account_id in self.accounts:
                self.accounts[account_id]["status"] = status
                self.accounts[account_id]["last_error"] = error
                self.save_accounts()

    # Common SMTP host corrections
    SMTP_HOST_FIXES = {
        "gmail": "smtp.gmail.com", "gmail.com": "smtp.gmail.com",
        "smtp.google.com": "smtp.gmail.com", "mail.gmail.com": "smtp.gmail.com",
        "outlook": "smtp-mail.outlook.com", "outlook.com": "smtp-mail.outlook.com",
        "smtp.outlook.com": "smtp-mail.outlook.com", "hotmail": "smtp-mail.outlook.com",
        "hotmail.com": "smtp-mail.outlook.com", "smtp.hotmail.com": "smtp-mail.outlook.com",
        "live.com": "smtp-mail.outlook.com", "smtp.live.com": "smtp-mail.outlook.com",
        "yahoo": "smtp.mail.yahoo.com", "yahoo.com": "smtp.mail.yahoo.com",
        "smtp.yahoo.com": "smtp.mail.yahoo.com",
        "zoho": "smtp.zoho.com", "zoho.com": "smtp.zoho.com",
        "icloud": "smtp.mail.me.com", "icloud.com": "smtp.mail.me.com",
        "aol": "smtp.aol.com", "aol.com": "smtp.aol.com",
    }

    def _suggest_smtp_host(self, bad_host: str) -> str:
        """Suggest correct SMTP host for common mistakes."""
        key = bad_host.lower().strip()
        if key in self.SMTP_HOST_FIXES:
            return self.SMTP_HOST_FIXES[key]
        # Check if email domain gives a hint
        for keyword, correct in self.SMTP_HOST_FIXES.items():
            if keyword in key:
                return correct
        return ""

    def test_account(self, account_id: str) -> dict:
        acc = self.get_account(account_id)
        if not acc:
            return {"success": False, "message": "Account not found", "account_id": account_id}

        try:
            host = acc["smtp_host"].strip()
            port = acc["smtp_port"]
            security = acc["smtp_security"]
            password = acc["password"]

            # Verify password was decrypted properly
            if not password or password.strip() == "":
                msg = "Password decryption failed. Delete this account and re-add it with your password."
                self.update_account_status(account_id, "failed", msg)
                return {"success": False, "message": msg, "account_id": account_id}

            # Strip ALL whitespace from app passwords (spaces, tabs, newlines)
            password_clean = "".join(password.split())

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

            server.login(acc["email"], password_clean)
            server.quit()

            self.update_account_status(account_id, "connected", None)
            return {"success": True, "message": f"✅ Connected to {host}:{port} as {acc['email']}", "account_id": account_id}

        except smtplib.SMTPAuthenticationError as e:
            error_msg = e.smtp_error.decode('utf-8', errors='replace') if isinstance(e.smtp_error, bytes) else str(e.smtp_error)
            msg = f"Authentication failed: {e.smtp_code} {error_msg}"
            self.update_account_status(account_id, "failed", msg)
            return {"success": False, "message": msg, "account_id": account_id}
        except smtplib.SMTPConnectError as e:
            msg = f"Connection refused by {acc['smtp_host']}:{acc['smtp_port']}. Check port and security settings."
            self.update_account_status(account_id, "failed", msg)
            return {"success": False, "message": msg, "account_id": account_id}
        except (OSError, socket.gaierror) as e:
            # DNS resolution failed — getaddrinfo error
            host = acc["smtp_host"]
            suggestion = self._suggest_smtp_host(host)
            msg = f"Cannot resolve hostname '{host}'. Check your internet or SMTP host."
            if suggestion and suggestion.lower() != host.lower():
                msg += f"\n💡 Did you mean: {suggestion}?"
            self.update_account_status(account_id, "failed", msg)
            return {"success": False, "message": msg, "account_id": account_id, "suggestion": suggestion}
        except (socket.timeout, TimeoutError):
            msg = f"Connection timed out to {acc['smtp_host']}:{acc['smtp_port']}. Server may be down or port blocked by firewall."
            self.update_account_status(account_id, "failed", msg)
            return {"success": False, "message": msg, "account_id": account_id}
        except ConnectionRefusedError:
            msg = f"Connection refused by {acc['smtp_host']}:{acc['smtp_port']}. Wrong port or server not accepting connections."
            self.update_account_status(account_id, "failed", msg)
            return {"success": False, "message": msg, "account_id": account_id}
        except Exception as e:
            msg = f"Connection error: {type(e).__name__} — {e}"
            self.update_account_status(account_id, "failed", msg)
            return {"success": False, "message": msg, "account_id": account_id}

    def test_all_accounts(self) -> list:
        results = []
        threads = []

        def _test(aid, container):
            res = self.test_account(aid)
            container.append(res)

        for aid in list(self.accounts.keys()):
            t = threading.Thread(target=_test, args=(aid, results), daemon=True)
            threads.append(t)
            t.start()

        for t in threads:
            t.join(timeout=60)

        return results

    # ─── Rotation ─────────────────────────────────────────────────

    def get_rotation_order(self) -> list:
        self.reset_daily_counters()
        eligible = []
        for aid, acc in self.accounts.items():
            remaining = acc["daily_limit"] - acc["sent_today"]
            if remaining > 0 and acc["status"] != "failed" and not acc.get("auto_disabled", False):
                health = acc.get("health_score", 100)
                eligible.append((aid, remaining, acc["status"], health))

        # Sort: connected first, then by health, then by capacity
        def sort_key(item):
            status_priority = 0 if item[2] == "connected" else 1
            return (status_priority, -item[3], -item[1])

        eligible.sort(key=sort_key)
        return [item[0] for item in eligible]

    # ─── Import / Export ──────────────────────────────────────────

    def export_accounts(self, output_path: str) -> bool:
        """Export all accounts to an encrypted backup file."""
        try:
            export_data = []
            for acc in self.accounts.values():
                entry = dict(acc)
                # Keep password encrypted
                export_data.append(entry)

            backup = {
                "version": "2.0",
                "exported_at": datetime.now().isoformat(),
                "account_count": len(export_data),
                "accounts": export_data,
            }
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(backup, f, indent=2, default=str)
            return True
        except Exception as e:
            print(f"[AccountManager] Export error: {e}")
            return False

    def import_accounts(self, input_path: str, overwrite: bool = False) -> dict:
        """Import accounts from a backup file."""
        try:
            with open(input_path, "r", encoding="utf-8") as f:
                backup = json.load(f)

            imported = 0
            skipped = 0
            for acc in backup.get("accounts", []):
                email = acc.get("email", "").lower()
                # Check if already exists
                exists = any(a["email"].lower() == email for a in self.accounts.values())
                if exists and not overwrite:
                    skipped += 1
                    continue

                # Remove old id if overwriting, generate new
                if exists and overwrite:
                    old_id = next(aid for aid, a in self.accounts.items() if a["email"].lower() == email)
                    del self.accounts[old_id]

                new_id = str(uuid.uuid4())[:8]
                acc["id"] = new_id
                with self._lock:
                    self.accounts[new_id] = acc
                imported += 1

            self.save_accounts()
            return {"success": True, "imported": imported, "skipped": skipped}
        except Exception as e:
            return {"success": False, "error": str(e), "imported": 0, "skipped": 0}

    # ─── Stats ────────────────────────────────────────────────────

    def get_stats(self) -> dict:
        total = len(self.accounts)
        active = sum(1 for a in self.accounts.values()
                     if a["status"] in ("connected", "unknown") and not a.get("auto_disabled", False))
        failed = sum(1 for a in self.accounts.values()
                     if a["status"] == "failed" or a.get("auto_disabled", False))
        limit_reached = sum(1 for a in self.accounts.values() if a["status"] == "limit_reached")
        total_capacity = sum(a["daily_limit"] for a in self.accounts.values())
        total_sent = sum(a["sent_today"] for a in self.accounts.values())
        avg_health = round(
            sum(a.get("health_score", 100) for a in self.accounts.values()) / max(total, 1), 1)
        lifetime_sent = sum(a.get("total_sent", 0) for a in self.accounts.values())
        lifetime_failed = sum(a.get("total_failed", 0) for a in self.accounts.values())
        return {
            "total_accounts": total,
            "active_accounts": active,
            "failed_accounts": failed,
            "limit_reached_accounts": limit_reached,
            "total_capacity_today": total_capacity,
            "total_sent_today": total_sent,
            "total_remaining_today": total_capacity - total_sent,
            "avg_health_score": avg_health,
            "lifetime_sent": lifetime_sent,
            "lifetime_failed": lifetime_failed,
        }

    # ─── Persistence ──────────────────────────────────────────────

    def save_accounts(self):
        try:
            data = json.dumps(self.accounts, indent=2, default=str)
            with open(self.accounts_path, "w", encoding="utf-8") as f:
                f.write(data)
        except IOError as e:
            print(f"[AccountManager] Error saving accounts: {e}")

    def load_accounts(self):
        if not self.accounts_path.exists():
            self.accounts = {}
            return
        try:
            with open(self.accounts_path, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if content:
                    self.accounts = json.loads(content)
                    # Migrate old accounts to v2 schema
                    for acc in self.accounts.values():
                        if "health_score" not in acc:
                            acc["health_score"] = 100.0
                        if "total_sent" not in acc:
                            acc["total_sent"] = 0
                        if "total_failed" not in acc:
                            acc["total_failed"] = 0
                        if "hard_bounces" not in acc:
                            acc["hard_bounces"] = 0
                        if "soft_bounces" not in acc:
                            acc["soft_bounces"] = 0
                        if "auto_disabled" not in acc:
                            acc["auto_disabled"] = False
                        if "last_send_time" not in acc:
                            acc["last_send_time"] = None
                        if "reply_to" not in acc:
                            acc["reply_to"] = acc.get("email", "")
                else:
                    self.accounts = {}
        except (json.JSONDecodeError, IOError) as e:
            print(f"[AccountManager] Error loading accounts, starting fresh: {e}")
            self.accounts = {}
