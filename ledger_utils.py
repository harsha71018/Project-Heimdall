"""Backward-compatible migration of persisted job URL identities."""
import shutil
import uuid

from url_utils import canonicalize_job_url


def backup_identity_migration(path):
    """Keep the original bytes before replacing keys; abort if backup fails."""
    backup = f"{path}.identity-{uuid.uuid4().hex}.bak"
    shutil.copy2(path, backup)
    return backup


def normalize_applications(ledger):
    normalized = {}
    for old_key, record in ledger.items():
        key = canonicalize_job_url(old_key) or old_key
        if key not in normalized:
            normalized[key] = dict(record)
            continue
        existing = normalized[key]
        # Retain both source records, including conflicting notes/status, for recovery.
        history = existing.setdefault("_identity_conflicts", [])
        if not history:
            history.append({"key": key, "record": {k: v for k, v in existing.items()
                                                   if k != "_identity_conflicts"}})
        history.append({"key": old_key, "record": record})
        field_times = dict(existing.get("_field_edited_at", {}))
        for field in ("status", "notes"):
            def rank(entry):
                timestamp = entry.get("_field_edited_at", {}).get(field, entry.get("_edited_at")) or 0
                meaningful = entry.get(field, "") not in ("", "Not Applied", None)
                return (timestamp, str(entry.get("last_updated", "")), meaningful)
            if rank(record) > rank(existing):
                existing[field] = record.get(field, "")
                field_times[field] = record.get("_field_edited_at", {}).get(field, record.get("_edited_at")) or 0
        # Preserve legacy timestamps on fields not replaced above.
        for field in ("status", "notes"):
            field_times.setdefault(field, existing.get("_edited_at") or 0)
        existing["_field_edited_at"] = field_times
        existing["_edited_at"] = max(field_times.values())
        for field, value in record.items():
            if field in ("status", "notes"):
                continue  # An intentional newer clear must not be filled from an older record.
            if field not in existing or existing[field] in (None, ""):
                existing[field] = value
    return normalized
