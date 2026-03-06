"""
Scheduler v1.0 — Campaign scheduling with timezone support.
Schedule campaigns for a specific date/time, with recurring options.
Persistent storage in schedules.json.
"""

import json
import time
import threading
import uuid
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, Optional, Dict, List


class CampaignScheduler:
    """
    Schedule email campaigns for future delivery.
    Supports one-time and recurring schedules.
    """

    SCHEDULE_FILE = "schedules.json"

    def __init__(self, data_dir: str = "."):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.schedule_path = self.data_dir / self.SCHEDULE_FILE
        self.schedules = {}
        self.active_timers = {}
        self._lock = threading.Lock()
        self.load_schedules()

    # ─── Schedule Management ──────────────────────────────────────

    def create_schedule(self, name: str, send_at: datetime,
                       campaign_config: dict,
                       recurrence: str = "none") -> dict:
        """
        Create a new scheduled campaign.
        recurrence: 'none', 'daily', 'weekly', 'monthly'
        """
        schedule_id = str(uuid.uuid4())[:8]
        schedule = {
            "id": schedule_id,
            "name": name.strip(),
            "send_at": send_at.isoformat(),
            "campaign_config": campaign_config,
            "recurrence": recurrence,
            "status": "scheduled",
            "created_at": datetime.now().isoformat(),
            "last_run": None,
            "run_count": 0,
            "next_run": send_at.isoformat(),
        }

        with self._lock:
            self.schedules[schedule_id] = schedule
            self.save_schedules()

        return schedule

    def cancel_schedule(self, schedule_id: str) -> bool:
        """Cancel a scheduled campaign."""
        with self._lock:
            if schedule_id not in self.schedules:
                return False
            self.schedules[schedule_id]["status"] = "cancelled"
            # Cancel timer if active
            timer = self.active_timers.pop(schedule_id, None)
            if timer:
                timer.cancel()
            self.save_schedules()
        return True

    def delete_schedule(self, schedule_id: str) -> bool:
        """Delete a schedule permanently."""
        with self._lock:
            if schedule_id not in self.schedules:
                return False
            timer = self.active_timers.pop(schedule_id, None)
            if timer:
                timer.cancel()
            del self.schedules[schedule_id]
            self.save_schedules()
        return True

    def get_all_schedules(self) -> List[dict]:
        """Get all schedules."""
        return list(self.schedules.values())

    def get_pending_schedules(self) -> List[dict]:
        """Get schedules that are pending execution."""
        return [s for s in self.schedules.values() if s["status"] == "scheduled"]

    def get_schedule(self, schedule_id: str) -> Optional[dict]:
        """Get a specific schedule."""
        return self.schedules.get(schedule_id)

    # ─── Timer Management ─────────────────────────────────────────

    def start_timer(self, schedule_id: str, callback: Callable):
        """Start a timer for a scheduled campaign."""
        schedule = self.schedules.get(schedule_id)
        if not schedule or schedule["status"] != "scheduled":
            return False

        send_at = datetime.fromisoformat(schedule["send_at"])
        now = datetime.now()
        delay_seconds = (send_at - now).total_seconds()

        if delay_seconds <= 0:
            # Execute immediately
            threading.Thread(target=self._execute, args=(schedule_id, callback), daemon=True).start()
            return True

        # Set timer
        timer = threading.Timer(delay_seconds, self._execute, args=(schedule_id, callback))
        timer.daemon = True
        timer.start()
        self.active_timers[schedule_id] = timer
        return True

    def start_all_timers(self, callback: Callable):
        """Start timers for all pending schedules."""
        for schedule in self.get_pending_schedules():
            self.start_timer(schedule["id"], callback)

    def _execute(self, schedule_id: str, callback: Callable):
        """Execute a scheduled campaign."""
        schedule = self.schedules.get(schedule_id)
        if not schedule or schedule["status"] != "scheduled":
            return

        with self._lock:
            schedule["status"] = "running"
            schedule["last_run"] = datetime.now().isoformat()
            schedule["run_count"] += 1
            self.save_schedules()

        try:
            callback(schedule["campaign_config"])
        except Exception as e:
            with self._lock:
                schedule["status"] = "failed"
                schedule["last_error"] = str(e)
                self.save_schedules()
            return

        with self._lock:
            # Handle recurrence
            if schedule["recurrence"] == "none":
                schedule["status"] = "completed"
            elif schedule["recurrence"] in ("daily", "weekly", "monthly"):
                send_at = datetime.fromisoformat(schedule["send_at"])
                if schedule["recurrence"] == "daily":
                    next_run = send_at + timedelta(days=1)
                elif schedule["recurrence"] == "weekly":
                    next_run = send_at + timedelta(weeks=1)
                else:  # monthly
                    next_run = send_at + timedelta(days=30)

                schedule["send_at"] = next_run.isoformat()
                schedule["next_run"] = next_run.isoformat()
                schedule["status"] = "scheduled"
                # Restart timer
                self.start_timer(schedule_id, callback)
            else:
                schedule["status"] = "completed"
            self.save_schedules()

    # ─── Time Utilities ───────────────────────────────────────────

    @staticmethod
    def get_time_until(send_at: datetime) -> str:
        """Get human-readable time until scheduled send."""
        now = datetime.now()
        delta = send_at - now

        if delta.total_seconds() <= 0:
            return "Now"

        days = delta.days
        hours, remainder = divmod(delta.seconds, 3600)
        minutes = remainder // 60

        parts = []
        if days > 0:
            parts.append(f"{days}d")
        if hours > 0:
            parts.append(f"{hours}h")
        if minutes > 0:
            parts.append(f"{minutes}m")

        return " ".join(parts) if parts else "<1m"

    @staticmethod
    def parse_schedule_time(date_str: str, time_str: str) -> datetime:
        """Parse date and time strings into datetime."""
        # Try multiple formats
        for fmt in ["%Y-%m-%d %H:%M", "%m/%d/%Y %H:%M", "%d/%m/%Y %H:%M",
                     "%Y-%m-%d %I:%M %p", "%m/%d/%Y %I:%M %p"]:
            try:
                return datetime.strptime(f"{date_str} {time_str}", fmt)
            except ValueError:
                continue
        raise ValueError(f"Cannot parse date/time: {date_str} {time_str}")

    # ─── Persistence ──────────────────────────────────────────────

    def save_schedules(self):
        try:
            with open(self.schedule_path, "w", encoding="utf-8") as f:
                json.dump(self.schedules, f, indent=2, default=str)
        except IOError as e:
            print(f"[Scheduler] Save error: {e}")

    def load_schedules(self):
        if not self.schedule_path.exists():
            self.schedules = {}
            return
        try:
            with open(self.schedule_path, "r", encoding="utf-8") as f:
                content = f.read().strip()
                self.schedules = json.loads(content) if content else {}
        except (json.JSONDecodeError, IOError):
            self.schedules = {}
