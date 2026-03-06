"""
ContactManager v1.0 — Contact segmentation and group management.
Save validated contact lists as named groups, tag contacts,
filter/search, merge/split groups. Persistent storage.
"""

import json
import uuid
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional


class ContactManager:
    """
    Manage contact groups with tagging, filtering, and segmentation.
    """

    CONTACTS_FILE = "contacts.json"

    # Default tags
    DEFAULT_TAGS = ["VIP", "New", "Cold", "Warm", "Hot", "Inactive", "Bounced"]

    def __init__(self, data_dir: str = "."):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.contacts_path = self.data_dir / self.CONTACTS_FILE
        self.groups = {}
        self.load_contacts()

    # ─── Group Management ─────────────────────────────────────────

    def create_group(self, name: str, contacts: list = None,
                     tags: list = None, description: str = "") -> dict:
        """Create a new contact group."""
        group_id = str(uuid.uuid4())[:8]
        group = {
            "id": group_id,
            "name": name.strip(),
            "description": description,
            "contacts": contacts or [],
            "tags": tags or [],
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat(),
            "count": len(contacts) if contacts else 0,
        }
        self.groups[group_id] = group
        self.save_contacts()
        return group

    def save_recipients_as_group(self, name: str, valid_emails: list,
                                  tags: list = None) -> dict:
        """Save a list of validated recipients as a named group."""
        contacts = []
        for r in valid_emails:
            contact = {
                "email": r.get("_email", ""),
                "name": r.get("name", "") or r.get("full_name", ""),
                "tags": tags or [],
                "data": {k: v for k, v in r.items() if not k.startswith("_")},
                "added_at": datetime.now().isoformat(),
            }
            contacts.append(contact)
        return self.create_group(name, contacts, tags)

    def delete_group(self, group_id: str) -> bool:
        if group_id in self.groups:
            del self.groups[group_id]
            self.save_contacts()
            return True
        return False

    def rename_group(self, group_id: str, new_name: str) -> bool:
        if group_id in self.groups:
            self.groups[group_id]["name"] = new_name.strip()
            self.groups[group_id]["updated_at"] = datetime.now().isoformat()
            self.save_contacts()
            return True
        return False

    def get_group(self, group_id: str) -> Optional[dict]:
        return self.groups.get(group_id)

    def get_group_by_name(self, name: str) -> Optional[dict]:
        for g in self.groups.values():
            if g["name"].lower() == name.lower():
                return g
        return None

    def get_all_groups(self) -> List[dict]:
        """Get summary of all groups (without full contact lists)."""
        summaries = []
        for g in self.groups.values():
            summaries.append({
                "id": g["id"],
                "name": g["name"],
                "description": g.get("description", ""),
                "count": len(g.get("contacts", [])),
                "tags": g.get("tags", []),
                "created_at": g["created_at"],
            })
        return summaries

    def get_group_contacts(self, group_id: str) -> list:
        """Get all contacts in a group, formatted for sending."""
        group = self.groups.get(group_id)
        if not group:
            return []
        result = []
        for c in group.get("contacts", []):
            recipient = dict(c.get("data", {}))
            recipient["_email"] = c.get("email", "")
            recipient["name"] = c.get("name", "")
            result.append(recipient)
        return result

    # ─── Tagging ──────────────────────────────────────────────────

    def tag_contacts(self, group_id: str, emails: list, tag: str):
        """Add a tag to specific contacts in a group."""
        group = self.groups.get(group_id)
        if not group:
            return
        for contact in group.get("contacts", []):
            if contact["email"] in emails:
                if tag not in contact.get("tags", []):
                    contact.setdefault("tags", []).append(tag)
        self.groups[group_id]["updated_at"] = datetime.now().isoformat()
        self.save_contacts()

    def untag_contacts(self, group_id: str, emails: list, tag: str):
        """Remove a tag from specific contacts."""
        group = self.groups.get(group_id)
        if not group:
            return
        for contact in group.get("contacts", []):
            if contact["email"] in emails and tag in contact.get("tags", []):
                contact["tags"].remove(tag)
        self.save_contacts()

    def filter_by_tag(self, group_id: str, tag: str) -> list:
        """Get contacts in a group that have a specific tag."""
        group = self.groups.get(group_id)
        if not group:
            return []
        return [c for c in group.get("contacts", []) if tag in c.get("tags", [])]

    # ─── Filtering & Search ───────────────────────────────────────

    def search_contacts(self, group_id: str, query: str) -> list:
        """Search contacts by email or name."""
        group = self.groups.get(group_id)
        if not group:
            return []
        query_lower = query.lower()
        return [c for c in group.get("contacts", [])
                if query_lower in c.get("email", "").lower()
                or query_lower in c.get("name", "").lower()]

    def filter_by_domain(self, group_id: str, domain: str) -> list:
        """Get contacts with a specific email domain."""
        group = self.groups.get(group_id)
        if not group:
            return []
        return [c for c in group.get("contacts", [])
                if c.get("email", "").lower().endswith(f"@{domain.lower()}")]

    # ─── Merge & Split ────────────────────────────────────────────

    def merge_groups(self, group_ids: list, new_name: str) -> dict:
        """Merge multiple groups into a new group, deduplicating by email."""
        all_contacts = []
        seen_emails = set()
        all_tags = set()

        for gid in group_ids:
            group = self.groups.get(gid)
            if not group:
                continue
            all_tags.update(group.get("tags", []))
            for contact in group.get("contacts", []):
                email = contact.get("email", "").lower()
                if email not in seen_emails:
                    seen_emails.add(email)
                    all_contacts.append(contact)

        return self.create_group(new_name, all_contacts, list(all_tags))

    def split_group(self, group_id: str, tag: str) -> tuple:
        """Split a group into two: contacts with tag and without."""
        group = self.groups.get(group_id)
        if not group:
            return None, None

        with_tag = [c for c in group.get("contacts", []) if tag in c.get("tags", [])]
        without_tag = [c for c in group.get("contacts", []) if tag not in c.get("tags", [])]

        g1 = self.create_group(f"{group['name']} [{tag}]", with_tag, [tag])
        g2 = self.create_group(f"{group['name']} [not {tag}]", without_tag)
        return g1, g2

    # ─── Stats ────────────────────────────────────────────────────

    def get_group_stats(self, group_id: str) -> dict:
        """Get statistics for a group."""
        group = self.groups.get(group_id)
        if not group:
            return {}

        contacts = group.get("contacts", [])
        domains = {}
        tags_count = {}

        for c in contacts:
            email = c.get("email", "")
            if "@" in email:
                domain = email.split("@")[1]
                domains[domain] = domains.get(domain, 0) + 1
            for t in c.get("tags", []):
                tags_count[t] = tags_count.get(t, 0) + 1

        return {
            "total_contacts": len(contacts),
            "domains": dict(sorted(domains.items(), key=lambda x: -x[1])[:10]),
            "tags": tags_count,
        }

    # ─── Persistence ──────────────────────────────────────────────

    def save_contacts(self):
        try:
            with open(self.contacts_path, "w", encoding="utf-8") as f:
                json.dump(self.groups, f, indent=2, default=str)
        except IOError as e:
            print(f"[ContactManager] Save error: {e}")

    def load_contacts(self):
        if not self.contacts_path.exists():
            self.groups = {}
            return
        try:
            with open(self.contacts_path, "r", encoding="utf-8") as f:
                content = f.read().strip()
                self.groups = json.loads(content) if content else {}
        except (json.JSONDecodeError, IOError):
            self.groups = {}
