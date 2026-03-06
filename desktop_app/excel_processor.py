"""
ExcelProcessor v2.0 — Advanced Excel/CSV loading and email validation.
Features: Multi-format support, auto-detect, typo correction, disposable email detection,
MX record validation, role-based filtering, smart name extraction, domain statistics.
"""

import pandas as pd
import re
import csv
import os
import socket
from typing import Optional, List, Dict
from pathlib import Path


class ExcelProcessor:
    """
    Loads Excel/CSV files, validates email addresses, corrects typos,
    detects disposable emails, validates MX records, filters role-based addresses,
    extracts personalization variables, and provides domain statistics.
    """

    EMAIL_REGEX = re.compile(
        r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$'
    )

    COMMON_TYPOS = {
        'gmial.com': 'gmail.com', 'gmai.com': 'gmail.com', 'gamil.com': 'gmail.com',
        'gmail.co': 'gmail.com', 'gmal.com': 'gmail.com', 'gnail.com': 'gmail.com',
        'gmail.con': 'gmail.com', 'gmail.cim': 'gmail.com', 'gmail.cm': 'gmail.com',
        'gmil.com': 'gmail.com', 'gmaill.com': 'gmail.com', 'gmail.comm': 'gmail.com',
        'yaho.com': 'yahoo.com', 'yahooo.com': 'yahoo.com', 'yahoo.co': 'yahoo.com',
        'yahoo.con': 'yahoo.com', 'yahho.com': 'yahoo.com', 'yhaoo.com': 'yahoo.com',
        'hotmial.com': 'hotmail.com', 'hotmai.com': 'hotmail.com', 'hotmail.co': 'hotmail.com',
        'hotamil.com': 'hotmail.com', 'hotmil.com': 'hotmail.com', 'hotmail.con': 'hotmail.com',
        'outlok.com': 'outlook.com', 'outloo.com': 'outlook.com', 'outlook.co': 'outlook.com',
        'outllok.com': 'outlook.com', 'outlool.com': 'outlook.com', 'outlook.con': 'outlook.com',
        'iclod.com': 'icloud.com', 'icloud.co': 'icloud.com',
        'protonmai.com': 'protonmail.com', 'protonmal.com': 'protonmail.com',
    }

    # Disposable/temporary email domains
    DISPOSABLE_DOMAINS = {
        'mailinator.com', 'guerrillamail.com', 'guerrillamail.net', 'tempmail.com',
        'throwaway.email', 'yopmail.com', 'temp-mail.org', 'fakeinbox.com',
        'sharklasers.com', 'guerrillamailblock.com', 'grr.la', 'dispostable.com',
        'trashmail.com', 'trashmail.net', 'trashmail.org', 'mailnesia.com',
        'maildrop.cc', 'discard.email', 'getnada.com', 'mohmal.com',
        'tempail.com', 'tempr.email', 'tempmailo.com', 'emailondeck.com',
        '10minutemail.com', 'throwaway.email', 'mintemail.com', 'mailsac.com',
        'harakirimail.com', 'burnermail.io', 'inboxkitten.com', 'trash-mail.com',
        'snapmail.cc', 'dropmail.me', 'tempmailaddress.com', 'crazymailing.com',
    }

    # Role-based email prefixes (often not personal inboxes)
    ROLE_PREFIXES = {
        'admin', 'info', 'noreply', 'no-reply', 'support', 'sales',
        'contact', 'webmaster', 'postmaster', 'abuse', 'help', 'billing',
        'marketing', 'newsletter', 'team', 'hello', 'office', 'mail',
        'service', 'feedback', 'enquiry', 'inquiry',
    }

    def __init__(self):
        self.df = None
        self.file_path = None
        self.email_column = None
        self.valid_emails = []
        self.invalid_emails = []
        self.duplicate_emails = []
        self.disposable_emails = []
        self.role_emails = []
        self.stats = {}
        self._mx_cache = {}  # domain → bool (has MX)

    def load(self, file_path: str) -> dict:
        """Load Excel or CSV file and return basic stats."""
        self.file_path = file_path
        ext = os.path.splitext(file_path)[1].lower()
        file_size = os.path.getsize(file_path)

        try:
            if ext in ('.xlsx', '.xls'):
                engine = 'openpyxl' if ext == '.xlsx' else 'xlrd'
                self.df = pd.read_excel(file_path, engine=engine)
            elif ext == '.csv':
                self.df = self._load_csv(file_path)
            else:
                return {
                    "success": False, "rows": 0, "columns": [],
                    "file_size_kb": round(file_size / 1024, 1),
                    "error": f"Unsupported file format: {ext}. Use .xlsx, .xls, or .csv"
                }

            # Clean up
            self.df.columns = [str(c).strip() for c in self.df.columns]
            for col in self.df.select_dtypes(include=['object']).columns:
                self.df[col] = self.df[col].apply(
                    lambda x: x.strip() if isinstance(x, str) else x
                )
            self.df.dropna(how='all', inplace=True)
            self.df.reset_index(drop=True, inplace=True)

            return {
                "success": True,
                "rows": len(self.df),
                "columns": list(self.df.columns),
                "file_size_kb": round(file_size / 1024, 1),
                "error": None
            }

        except Exception as e:
            return {
                "success": False, "rows": 0, "columns": [],
                "file_size_kb": round(file_size / 1024, 1),
                "error": str(e)
            }

    def _load_csv(self, file_path: str) -> pd.DataFrame:
        """Try multiple encodings and delimiters for CSV files."""
        encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'iso-8859-1']
        delimiters = [',', ';', '\t', '|']
        for enc in encodings:
            for delim in delimiters:
                try:
                    df = pd.read_csv(file_path, encoding=enc, sep=delim)
                    if len(df.columns) > 1 or len(df) > 0:
                        return df
                except (UnicodeDecodeError, UnicodeError, pd.errors.ParserError):
                    continue
        # Fallback: try with default
        return pd.read_csv(file_path, encoding='latin-1')

    def auto_detect_email_column(self) -> Optional[str]:
        """Score each column by how many values match email regex. Return best match."""
        if self.df is None or self.df.empty:
            return None

        best_col = None
        best_score = 0.0

        for col in self.df.columns:
            try:
                # First check column name
                col_lower = col.lower().replace(' ', '')
                if col_lower in ('email', 'emailaddress', 'e-mail', 'mail', 'emailid'):
                    values = self.df[col].dropna().astype(str)
                    matches = sum(1 for v in values if self.EMAIL_REGEX.match(v.strip().lower()))
                    if matches > 0:
                        return col  # Exact name match with valid emails = best

                values = self.df[col].dropna().astype(str)
                if len(values) == 0:
                    continue
                matches = sum(1 for v in values if self.EMAIL_REGEX.match(v.strip().lower()))
                score = matches / len(values)
                if score > best_score:
                    best_score = score
                    best_col = col
            except Exception:
                continue

        if best_score >= 0.3:
            self.email_column = best_col
            return best_col
        return None

    def fix_typos(self, email: str) -> str:
        """Correct common domain typos in email addresses."""
        if '@' not in email:
            return email
        local, domain = email.rsplit('@', 1)
        domain = domain.lower().strip('.')
        corrected_domain = self.COMMON_TYPOS.get(domain, domain)
        return f"{local}@{corrected_domain}"

    def is_disposable(self, email: str) -> bool:
        """Check if email uses a disposable/temporary domain."""
        if '@' not in email:
            return False
        domain = email.split('@')[1].lower()
        return domain in self.DISPOSABLE_DOMAINS

    def is_role_email(self, email: str) -> bool:
        """Check if email is a role-based address (admin@, info@, etc.)."""
        if '@' not in email:
            return False
        local = email.split('@')[0].lower()
        return local in self.ROLE_PREFIXES

    def check_mx_record(self, domain: str) -> bool:
        """Check if domain has MX records (valid mail server)."""
        if domain in self._mx_cache:
            return self._mx_cache[domain]

        try:
            import dns.resolver
            answers = dns.resolver.resolve(domain, 'MX')
            has_mx = len(answers) > 0
        except ImportError:
            # dns.resolver not available, try socket fallback
            try:
                socket.getaddrinfo(domain, 25, socket.AF_INET)
                has_mx = True
            except socket.gaierror:
                has_mx = False
        except Exception:
            has_mx = False

        self._mx_cache[domain] = has_mx
        return has_mx

    def extract_name_from_email(self, email: str) -> str:
        """Try to extract a name from email address (john.smith@ → John Smith)."""
        if '@' not in email:
            return ""
        local = email.split('@')[0].lower()
        # Remove common prefixes/suffixes
        local = re.sub(r'\d+$', '', local)  # Remove trailing numbers
        # Split on dots, underscores, hyphens
        parts = re.split(r'[._\-]', local)
        parts = [p.capitalize() for p in parts if len(p) > 1 and not p.isdigit()]
        return ' '.join(parts[:3]) if parts else ""

    def validate_and_load(self, column: str, check_mx: bool = False,
                          filter_disposable: bool = True,
                          filter_roles: bool = False) -> dict:
        """
        Validate emails from the specified column with advanced filtering.
        Build valid_emails list with optional MX, disposable, and role filtering.
        """
        if self.df is None:
            return {"error": "No file loaded"}

        if column not in self.df.columns:
            return {"error": f"Column '{column}' not found in file"}

        self.email_column = column
        self.valid_emails = []
        self.invalid_emails = []
        self.duplicate_emails = []
        self.disposable_emails = []
        self.role_emails = []
        seen = set()
        empty_count = 0
        typo_fixed = 0
        mx_failed = 0

        for idx, row in self.df.iterrows():
            raw = row.get(column)
            if pd.isna(raw) or str(raw).strip() == '':
                empty_count += 1
                continue

            email = str(raw).strip().lower()
            original = email
            email = self.fix_typos(email)
            if email != original:
                typo_fixed += 1

            if not self.EMAIL_REGEX.match(email):
                self.invalid_emails.append({
                    "_email": original,
                    "_row_index": idx + 2,
                    "_reason": "Invalid format"
                })
                continue

            if email in seen:
                self.duplicate_emails.append({
                    "_email": email,
                    "_row_index": idx + 2,
                    "_reason": "Duplicate"
                })
                continue

            # Disposable check
            if filter_disposable and self.is_disposable(email):
                self.disposable_emails.append({
                    "_email": email,
                    "_row_index": idx + 2,
                    "_reason": "Disposable email domain"
                })
                continue

            # Role-based check
            if filter_roles and self.is_role_email(email):
                self.role_emails.append({
                    "_email": email,
                    "_row_index": idx + 2,
                    "_reason": "Role-based address"
                })
                continue

            # MX record check
            if check_mx:
                domain = email.split('@')[1]
                if not self.check_mx_record(domain):
                    self.invalid_emails.append({
                        "_email": email,
                        "_row_index": idx + 2,
                        "_reason": "No MX record (invalid domain)"
                    })
                    mx_failed += 1
                    continue

            seen.add(email)

            # Build recipient dict with all columns
            recipient = {}
            for col in self.df.columns:
                val = row.get(col)
                recipient[col.lower()] = str(val) if not pd.isna(val) else ""
            recipient["_email"] = email
            recipient["_row_index"] = idx + 2

            # Try to extract name if no name column exists
            name_cols = [c for c in self.df.columns if c.lower() in ('name', 'full_name', 'fullname', 'first_name')]
            if not name_cols:
                extracted = self.extract_name_from_email(email)
                if extracted:
                    recipient["name"] = extracted

            self.valid_emails.append(recipient)

        self.stats = {
            "total_rows": len(self.df),
            "valid": len(self.valid_emails),
            "invalid": len(self.invalid_emails),
            "duplicates": len(self.duplicate_emails),
            "disposable": len(self.disposable_emails),
            "role_based": len(self.role_emails),
            "empty": empty_count,
            "typos_fixed": typo_fixed,
            "mx_failed": mx_failed,
        }
        return self.stats

    def get_preview(self, n: int = 100) -> list:
        """Return first n rows as list of dicts for preview display."""
        if self.df is None:
            return []
        preview_df = self.df.head(n)
        return preview_df.fillna("").to_dict(orient='records')

    def get_columns(self) -> list:
        """Return list of column names."""
        if self.df is None:
            return []
        return list(self.df.columns)

    def get_personalization_vars(self) -> list:
        """Return list of {{variable}} strings based on column names."""
        if self.df is None:
            return []
        vars_list = []
        for col in self.df.columns:
            var_name = col.strip().lower().replace(' ', '_')
            vars_list.append("{{" + var_name + "}}")
        builtins = ["{{first_name}}", "{{last_name}}", "{{date}}", "{{time}}", "{{year}}", "{{unsubscribe_link}}"]
        for b in builtins:
            if b not in vars_list:
                vars_list.append(b)
        return vars_list

    def get_domain_stats(self) -> dict:
        """Get email domain distribution from valid emails."""
        domains = {}
        for recip in self.valid_emails:
            email = recip.get("_email", "")
            if "@" in email:
                domain = email.split("@")[1]
                domains[domain] = domains.get(domain, 0) + 1
        # Sort by count descending
        sorted_domains = dict(sorted(domains.items(), key=lambda x: -x[1]))
        total = len(self.valid_emails)
        stats = {}
        for domain, count in list(sorted_domains.items())[:15]:
            pct = round(count / max(total, 1) * 100, 1)
            stats[domain] = {"count": count, "percent": pct}
        return stats

    def export_invalid_emails(self, output_path: str) -> bool:
        """Write all problematic emails to CSV for review."""
        try:
            all_issues = (self.invalid_emails + self.duplicate_emails +
                         self.disposable_emails + self.role_emails)
            if not all_issues:
                return False
            with open(output_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['_email', '_row_index', '_reason'])
                writer.writeheader()
                writer.writerows(all_issues)
            return True
        except IOError as e:
            print(f"[ExcelProcessor] Export error: {e}")
            return False

    def get_recipient_count(self) -> int:
        """Return count of valid emails ready to send."""
        return len(self.valid_emails)
