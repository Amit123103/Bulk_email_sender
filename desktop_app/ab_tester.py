"""
ABTester v1.0 — A/B split testing for email campaigns.
Split recipients into groups, test different subject lines or bodies,
track results per variant, and declare winner.
"""

import random
import json
import uuid
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional


class ABTester:
    """
    A/B split testing engine for email campaigns.
    Supports testing different subject lines, email bodies, or both.
    """

    RESULTS_FILE = "ab_results.json"

    def __init__(self, data_dir: str = "."):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.results_path = self.data_dir / self.RESULTS_FILE
        self.tests = {}
        self.load_results()

    def create_test(self, name: str, variant_a: dict, variant_b: dict,
                    split_ratio: float = 0.5) -> dict:
        """
        Create an A/B test.
        variant_a/b: {"subject": str, "body": str}
        split_ratio: fraction of recipients for variant A (0.0-1.0)
        """
        test_id = str(uuid.uuid4())[:8]
        test = {
            "id": test_id,
            "name": name.strip(),
            "variant_a": variant_a,
            "variant_b": variant_b,
            "split_ratio": max(0.1, min(0.9, split_ratio)),
            "status": "created",
            "created_at": datetime.now().isoformat(),
            "results_a": {"sent": 0, "failed": 0, "bounced": 0, "recipients": []},
            "results_b": {"sent": 0, "failed": 0, "bounced": 0, "recipients": []},
            "winner": None,
        }
        self.tests[test_id] = test
        self.save_results()
        return test

    def split_recipients(self, test_id: str, recipients: list) -> tuple:
        """
        Split recipients into A and B groups based on ratio.
        Returns (group_a, group_b) lists.
        """
        test = self.tests.get(test_id)
        if not test:
            return recipients, []

        shuffled = list(recipients)
        random.shuffle(shuffled)

        split_point = int(len(shuffled) * test["split_ratio"])
        group_a = shuffled[:split_point]
        group_b = shuffled[split_point:]

        # Tag recipients with their variant
        for r in group_a:
            r["_ab_variant"] = "A"
        for r in group_b:
            r["_ab_variant"] = "B"

        return group_a, group_b

    def get_variant_config(self, test_id: str, variant: str) -> dict:
        """Get subject and body for a variant."""
        test = self.tests.get(test_id)
        if not test:
            return {}
        if variant == "A":
            return test["variant_a"]
        return test["variant_b"]

    def record_result(self, test_id: str, variant: str, email: str,
                      status: str, bounce_type: str = None):
        """Record a send result for the test."""
        test = self.tests.get(test_id)
        if not test:
            return

        key = "results_a" if variant == "A" else "results_b"
        if status == "sent":
            test[key]["sent"] += 1
        elif status == "failed":
            test[key]["failed"] += 1
            if bounce_type:
                test[key]["bounced"] += 1
        test[key]["recipients"].append(email)
        self.save_results()

    def get_test_results(self, test_id: str) -> Optional[dict]:
        """Get results for an A/B test."""
        test = self.tests.get(test_id)
        if not test:
            return None

        ra = test["results_a"]
        rb = test["results_b"]

        total_a = ra["sent"] + ra["failed"]
        total_b = rb["sent"] + rb["failed"]

        rate_a = round(ra["sent"] / max(total_a, 1) * 100, 1)
        rate_b = round(rb["sent"] / max(total_b, 1) * 100, 1)

        # Determine winner
        if total_a >= 5 and total_b >= 5:
            if rate_a > rate_b + 2:
                winner = "A"
            elif rate_b > rate_a + 2:
                winner = "B"
            else:
                winner = "Tie"
        else:
            winner = "Insufficient data"

        test["winner"] = winner

        return {
            "test_id": test_id,
            "name": test["name"],
            "variant_a": {
                "subject": test["variant_a"].get("subject", ""),
                "sent": ra["sent"],
                "failed": ra["failed"],
                "total": total_a,
                "rate": rate_a,
            },
            "variant_b": {
                "subject": test["variant_b"].get("subject", ""),
                "sent": rb["sent"],
                "failed": rb["failed"],
                "total": total_b,
                "rate": rate_b,
            },
            "winner": winner,
        }

    def get_all_tests(self) -> List[dict]:
        """Get summary of all tests."""
        summaries = []
        for test in self.tests.values():
            ra = test["results_a"]
            rb = test["results_b"]
            summaries.append({
                "id": test["id"],
                "name": test["name"],
                "status": test["status"],
                "a_rate": round(ra["sent"] / max(ra["sent"] + ra["failed"], 1) * 100, 1),
                "b_rate": round(rb["sent"] / max(rb["sent"] + rb["failed"], 1) * 100, 1),
                "winner": test.get("winner"),
                "created_at": test["created_at"],
            })
        return summaries

    def delete_test(self, test_id: str) -> bool:
        if test_id in self.tests:
            del self.tests[test_id]
            self.save_results()
            return True
        return False

    # ─── Persistence ──────────────────────────────────────────────

    def save_results(self):
        try:
            with open(self.results_path, "w", encoding="utf-8") as f:
                json.dump(self.tests, f, indent=2, default=str)
        except IOError as e:
            print(f"[ABTester] Save error: {e}")

    def load_results(self):
        if not self.results_path.exists():
            self.tests = {}
            return
        try:
            with open(self.results_path, "r", encoding="utf-8") as f:
                content = f.read().strip()
                self.tests = json.loads(content) if content else {}
        except (json.JSONDecodeError, IOError):
            self.tests = {}
