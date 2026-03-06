"""
SpamChecker v1.0 — Email content spam score analyzer.
Analyzes email subject and body for 50+ spam trigger words, HTML ratio,
link density, ALL CAPS, exclamation marks, and missing best practices.
Returns a score 0-100 with detailed breakdown.
"""

import re
from typing import Dict, List, Tuple


class SpamChecker:
    """
    Analyze email content for spam indicators.
    Score 0-100: 0=clean, 100=definite spam.
    """

    # Spam trigger words/phrases grouped by severity
    HIGH_RISK_WORDS = [
        'act now', 'buy now', 'click here', 'click below', 'limited time',
        'urgent', 'free', 'winner', 'congratulations', 'you have been selected',
        'dear friend', 'no obligation', 'risk free', 'risk-free', 'double your',
        'earn money', 'make money', 'extra income', 'work from home',
        'credit card', 'no credit check', 'as seen on', 'info you requested',
        'this is not spam', 'not junk', 'not spam', 'remove me',
        'unsubscribe', 'opt out', 'opt-out',
    ]

    MEDIUM_RISK_WORDS = [
        'offer', 'discount', 'deal', 'sale', 'save', 'cheap', 'lowest price',
        'order now', 'buy direct', 'special promotion', 'limited offer',
        'exclusive deal', 'best price', 'bargain', 'bonus', 'prize',
        'guarantee', 'guaranteed', 'no cost', 'no fees', 'obligation',
        'cash', 'dollars', 'money back', 'refund', 'investment',
        'income', 'profit', 'earn', 'opportunity', 'no catch',
        'apply now', 'sign up free', 'join now', 'subscribe now',
        'act immediately', 'don\'t miss', 'don\'t delete', 'please read',
        'important information', 'verification required', 'confirm your',
        'verify your', 'update your', 'suspend', 'suspended',
    ]

    LOW_RISK_WORDS = [
        'free trial', 'gift', 'reward', 'claim', 'instant',
        'amazing', 'incredible', 'revolutionary', 'breakthrough',
        'miracle', 'secret', 'shocking', 'unbelievable',
        'once in a lifetime', 'time limited', 'while supplies last',
        'call now', 'reply now', 'respond now', 'contact us immediately',
    ]

    # Phishing indicators
    PHISHING_PATTERNS = [
        r'verify\s+your\s+(account|identity|email|password)',
        r'confirm\s+your\s+(account|identity|email|password)',
        r'update\s+your\s+(account|payment|billing)',
        r'your\s+account\s+(has been|will be|is)\s+(suspended|locked|closed)',
        r'click\s+(here|below)\s+to\s+(verify|confirm|update)',
        r'(unusual|suspicious)\s+(activity|login|sign)',
    ]

    def __init__(self):
        self.results = {}

    def analyze(self, subject: str, body: str, is_html: bool = True) -> Dict:
        """
        Full spam analysis. Returns dict with score and breakdown.
        """
        checks = []
        total_penalty = 0

        # 1. Subject line analysis
        subj_penalty, subj_details = self._check_subject(subject)
        checks.extend(subj_details)
        total_penalty += subj_penalty

        # 2. Spam trigger words
        word_penalty, word_details = self._check_trigger_words(subject + " " + body)
        checks.extend(word_details)
        total_penalty += word_penalty

        # 3. ALL CAPS analysis
        caps_penalty, caps_details = self._check_caps(subject, body)
        checks.extend(caps_details)
        total_penalty += caps_penalty

        # 4. Exclamation marks
        excl_penalty, excl_details = self._check_exclamation(subject, body)
        checks.extend(excl_details)
        total_penalty += excl_penalty

        # 5. HTML analysis
        if is_html:
            html_penalty, html_details = self._check_html(body)
            checks.extend(html_details)
            total_penalty += html_penalty

        # 6. Link analysis
        link_penalty, link_details = self._check_links(body)
        checks.extend(link_details)
        total_penalty += link_penalty

        # 7. Phishing patterns
        phish_penalty, phish_details = self._check_phishing(subject + " " + body)
        checks.extend(phish_details)
        total_penalty += phish_penalty

        # 8. Best practices
        bp_penalty, bp_details = self._check_best_practices(subject, body, is_html)
        checks.extend(bp_details)
        total_penalty += bp_penalty

        # Clamp score
        score = min(100, max(0, total_penalty))

        # Rating
        if score <= 15:
            rating = "EXCELLENT"
            color = "green"
            emoji = "🟢"
        elif score <= 30:
            rating = "GOOD"
            color = "green"
            emoji = "🟢"
        elif score <= 50:
            rating = "FAIR"
            color = "yellow"
            emoji = "🟡"
        elif score <= 70:
            rating = "POOR"
            color = "orange"
            emoji = "🟠"
        else:
            rating = "SPAM"
            color = "red"
            emoji = "🔴"

        self.results = {
            "score": score,
            "rating": rating,
            "color": color,
            "emoji": emoji,
            "checks": checks,
            "summary": f"{emoji} Score: {score}/100 — {rating}",
            "pass_count": sum(1 for c in checks if c["status"] == "pass"),
            "warn_count": sum(1 for c in checks if c["status"] == "warn"),
            "fail_count": sum(1 for c in checks if c["status"] == "fail"),
        }
        return self.results

    def _check_subject(self, subject: str) -> Tuple[int, List]:
        """Analyze subject line."""
        penalty = 0
        details = []

        if not subject or not subject.strip():
            penalty += 15
            details.append({"check": "Subject line", "status": "fail",
                           "message": "Missing subject line (+15)", "penalty": 15})
            return penalty, details

        if len(subject) > 78:
            penalty += 5
            details.append({"check": "Subject length", "status": "warn",
                           "message": f"Too long ({len(subject)} chars, max 78) (+5)", "penalty": 5})
        elif len(subject) < 10:
            penalty += 3
            details.append({"check": "Subject length", "status": "warn",
                           "message": f"Very short ({len(subject)} chars) (+3)", "penalty": 3})
        else:
            details.append({"check": "Subject length", "status": "pass",
                           "message": f"Good length ({len(subject)} chars)", "penalty": 0})

        if subject == subject.upper() and len(subject) > 5:
            penalty += 10
            details.append({"check": "Subject ALL CAPS", "status": "fail",
                           "message": "ALL CAPS subject (+10)", "penalty": 10})

        if subject.startswith("RE:") or subject.startswith("FW:"):
            penalty += 8
            details.append({"check": "Fake reply/forward", "status": "fail",
                           "message": "Fake RE:/FW: prefix (+8)", "penalty": 8})

        excl_count = subject.count('!')
        if excl_count > 2:
            penalty += 5
            details.append({"check": "Subject exclamation", "status": "warn",
                           "message": f"{excl_count} exclamation marks in subject (+5)", "penalty": 5})

        if '$' in subject or '€' in subject or '£' in subject:
            penalty += 5
            details.append({"check": "Currency in subject", "status": "warn",
                           "message": "Currency symbol in subject (+5)", "penalty": 5})

        return penalty, details

    def _check_trigger_words(self, text: str) -> Tuple[int, List]:
        """Check for spam trigger words."""
        penalty = 0
        details = []
        text_lower = text.lower()
        found_high = [w for w in self.HIGH_RISK_WORDS if w in text_lower]
        found_med = [w for w in self.MEDIUM_RISK_WORDS if w in text_lower]
        found_low = [w for w in self.LOW_RISK_WORDS if w in text_lower]

        if found_high:
            p = min(25, len(found_high) * 5)
            penalty += p
            details.append({"check": "High-risk spam words", "status": "fail",
                           "message": f"Found: {', '.join(found_high[:5])} (+{p})", "penalty": p})
        if found_med:
            p = min(15, len(found_med) * 2)
            penalty += p
            details.append({"check": "Medium-risk words", "status": "warn",
                           "message": f"Found: {', '.join(found_med[:5])} (+{p})", "penalty": p})
        if found_low:
            p = min(8, len(found_low) * 1)
            penalty += p
            details.append({"check": "Low-risk words", "status": "warn",
                           "message": f"Found: {', '.join(found_low[:5])} (+{p})", "penalty": p})
        if not (found_high or found_med or found_low):
            details.append({"check": "Trigger words", "status": "pass",
                           "message": "No spam trigger words found", "penalty": 0})
        return penalty, details

    def _check_caps(self, subject: str, body: str) -> Tuple[int, List]:
        """Check ALL CAPS percentage."""
        penalty = 0
        details = []
        text = subject + " " + body
        words = re.findall(r'[A-Za-z]+', text)
        if words:
            caps_words = [w for w in words if w == w.upper() and len(w) > 2]
            ratio = len(caps_words) / len(words)
            if ratio > 0.3:
                penalty += 10
                details.append({"check": "ALL CAPS ratio", "status": "fail",
                               "message": f"{ratio:.0%} of words in ALL CAPS (+10)", "penalty": 10})
            elif ratio > 0.15:
                penalty += 5
                details.append({"check": "ALL CAPS ratio", "status": "warn",
                               "message": f"{ratio:.0%} of words in ALL CAPS (+5)", "penalty": 5})
            else:
                details.append({"check": "ALL CAPS ratio", "status": "pass",
                               "message": f"{ratio:.0%} caps words — acceptable", "penalty": 0})
        return penalty, details

    def _check_exclamation(self, subject: str, body: str) -> Tuple[int, List]:
        """Check exclamation mark density."""
        penalty = 0
        details = []
        text = subject + " " + body
        excl = text.count('!')
        words = len(text.split())
        if words > 0:
            ratio = excl / words
            if ratio > 0.1:
                penalty += 8
                details.append({"check": "Exclamation density", "status": "fail",
                               "message": f"{excl} exclamation marks ({ratio:.1%} ratio) (+8)", "penalty": 8})
            elif ratio > 0.05:
                penalty += 4
                details.append({"check": "Exclamation density", "status": "warn",
                               "message": f"{excl} exclamation marks (+4)", "penalty": 4})
            else:
                details.append({"check": "Exclamation marks", "status": "pass",
                               "message": f"{excl} — acceptable", "penalty": 0})
        return penalty, details

    def _check_html(self, body: str) -> Tuple[int, List]:
        """Check HTML quality and ratio."""
        penalty = 0
        details = []
        plain = re.sub(r'<[^>]+>', '', body)
        html_len = len(body)
        text_len = len(plain.strip())

        if html_len > 0:
            ratio = text_len / html_len
            if ratio < 0.2:
                penalty += 8
                details.append({"check": "HTML/text ratio", "status": "warn",
                               "message": f"Low text ratio ({ratio:.0%}) — too much HTML (+8)", "penalty": 8})
            else:
                details.append({"check": "HTML/text ratio", "status": "pass",
                               "message": f"Text ratio: {ratio:.0%} — good", "penalty": 0})

        # Check for hidden text (font-size:0, display:none)
        if re.search(r'font-size\s*:\s*0', body) or re.search(r'display\s*:\s*none', body):
            penalty += 15
            details.append({"check": "Hidden text", "status": "fail",
                           "message": "Contains hidden text (font-size:0 or display:none) (+15)", "penalty": 15})

        # Check for embedded images
        img_count = len(re.findall(r'<img\s', body, re.I))
        if img_count > 5:
            penalty += 5
            details.append({"check": "Image count", "status": "warn",
                           "message": f"{img_count} images — may be flagged (+5)", "penalty": 5})

        return penalty, details

    def _check_links(self, body: str) -> Tuple[int, List]:
        """Check link count and quality."""
        penalty = 0
        details = []
        urls = re.findall(r'https?://\S+', body)
        href_links = re.findall(r'href\s*=\s*["\']([^"\']+)', body)
        all_links = urls + href_links

        if len(all_links) > 10:
            penalty += 5
            details.append({"check": "Link count", "status": "warn",
                           "message": f"{len(all_links)} links — high density (+5)", "penalty": 5})

        # Check for URL shorteners
        shorteners = ['bit.ly', 'tinyurl', 'goo.gl', 't.co', 'ow.ly', 'is.gd']
        found_short = [u for u in all_links if any(s in u.lower() for s in shorteners)]
        if found_short:
            penalty += 8
            details.append({"check": "URL shorteners", "status": "fail",
                           "message": f"Found shortened URLs (+8)", "penalty": 8})

        # Check for IP-based URLs
        ip_urls = [u for u in all_links if re.search(r'https?://\d{1,3}\.\d{1,3}\.', u)]
        if ip_urls:
            penalty += 10
            details.append({"check": "IP-based URLs", "status": "fail",
                           "message": "URLs use IP addresses instead of domains (+10)", "penalty": 10})

        if not all_links:
            details.append({"check": "Links", "status": "pass",
                           "message": "No links — clean", "penalty": 0})
        elif not (found_short or ip_urls) and len(all_links) <= 10:
            details.append({"check": "Links", "status": "pass",
                           "message": f"{len(all_links)} links — acceptable", "penalty": 0})

        return penalty, details

    def _check_phishing(self, text: str) -> Tuple[int, List]:
        """Check for phishing patterns."""
        penalty = 0
        details = []
        text_lower = text.lower()
        found = []
        for pattern in self.PHISHING_PATTERNS:
            if re.search(pattern, text_lower):
                found.append(pattern.split(r'\s+')[0].replace('\\', ''))

        if found:
            penalty += 15
            details.append({"check": "Phishing patterns", "status": "fail",
                           "message": f"Detected phishing-like language (+15)", "penalty": 15})
        else:
            details.append({"check": "Phishing check", "status": "pass",
                           "message": "No phishing patterns detected", "penalty": 0})
        return penalty, details

    def _check_best_practices(self, subject: str, body: str, is_html: bool) -> Tuple[int, List]:
        """Check email best practices."""
        penalty = 0
        details = []
        text_lower = body.lower()

        # Check for unsubscribe
        has_unsub = ('unsubscribe' in text_lower or 'opt out' in text_lower
                     or 'opt-out' in text_lower or '{{unsubscribe_link}}' in text_lower)
        if not has_unsub:
            penalty += 5
            details.append({"check": "Unsubscribe link", "status": "warn",
                           "message": "No unsubscribe option found (+5)", "penalty": 5})
        else:
            details.append({"check": "Unsubscribe link", "status": "pass",
                           "message": "Unsubscribe option present", "penalty": 0})

        # Check email length
        word_count = len(body.split())
        if word_count < 20:
            penalty += 3
            details.append({"check": "Content length", "status": "warn",
                           "message": f"Very short ({word_count} words) — may be flagged (+3)", "penalty": 3})
        elif word_count > 3000:
            penalty += 3
            details.append({"check": "Content length", "status": "warn",
                           "message": f"Very long ({word_count} words) (+3)", "penalty": 3})
        else:
            details.append({"check": "Content length", "status": "pass",
                           "message": f"{word_count} words — good", "penalty": 0})

        # Personalization check
        has_personalization = '{{' in body and '}}' in body
        if has_personalization:
            details.append({"check": "Personalization", "status": "pass",
                           "message": "Uses personalization variables", "penalty": 0})
        else:
            penalty += 2
            details.append({"check": "Personalization", "status": "warn",
                           "message": "No personalization — generic emails flag spam (+2)", "penalty": 2})

        return penalty, details

    def get_score(self) -> int:
        """Return last calculated score."""
        return self.results.get("score", 0)

    def get_summary(self) -> str:
        """Return human-readable summary."""
        return self.results.get("summary", "Not analyzed")

    def get_tips(self) -> List[str]:
        """Return actionable tips to improve score."""
        tips = []
        for check in self.results.get("checks", []):
            if check["status"] in ("warn", "fail"):
                tips.append(f"• {check['check']}: {check['message']}")
        if not tips:
            tips.append("• Your email looks great! No improvements needed.")
        return tips
