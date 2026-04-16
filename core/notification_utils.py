"""
File-backed in-app notifications for CMMS/PTW workflow events.
"""
from __future__ import annotations

import json
import uuid
from datetime import datetime
from pathlib import Path

from django.conf import settings

BASE_DIR = Path(
    getattr(settings, 'BASE_DIR', Path(__file__).resolve().parent.parent)
).resolve()
CMMS_DATA_DIR = Path(
    getattr(settings, 'CMMS_DATA_DIR', BASE_DIR / 'cmms_data')
).resolve()
NOTIFICATIONS_FILE = CMMS_DATA_DIR / 'notifications.json'

CMMS_DATA_DIR.mkdir(parents=True, exist_ok=True)


def _load() -> list[dict]:
    if not NOTIFICATIONS_FILE.exists():
        return []
    try:
        with NOTIFICATIONS_FILE.open('r', encoding='utf-8') as handle:
            data = json.load(handle)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _save(data: list[dict]) -> None:
    NOTIFICATIONS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with NOTIFICATIONS_FILE.open('w', encoding='utf-8') as handle:
        json.dump(data, handle, indent=2, ensure_ascii=False)


def _username(value: str | None) -> str:
    return str(value or '').strip().lower()


def _now() -> str:
    return datetime.now().isoformat()


def _sort_key(notification: dict) -> tuple:
    return (
        str(notification.get('created_at') or ''),
        str(notification.get('id') or ''),
    )


def _normalize_notification(notification: dict, username: str) -> dict:
    read_at = str(notification.get('read_at') or '').strip()
    return {
        'id': str(notification.get('id') or ''),
        'username': _username(notification.get('username')),
        'title': str(notification.get('title') or '').strip(),
        'message': str(notification.get('message') or '').strip(),
        'link': str(notification.get('link') or '').strip(),
        'kind': str(notification.get('kind') or 'info').strip() or 'info',
        'entity_type': str(notification.get('entity_type') or '').strip(),
        'entity_id': str(notification.get('entity_id') or '').strip(),
        'permit_id': str(notification.get('permit_id') or '').strip(),
        'actor_name': str(notification.get('actor_name') or '').strip(),
        'created_at': str(notification.get('created_at') or ''),
        'read_at': read_at,
        'is_unread': not bool(read_at),
        'is_mine': _username(notification.get('username')) == _username(username),
    }


def create_notification(
    username: str,
    *,
    title: str,
    message: str = '',
    link: str = '',
    kind: str = 'info',
    entity_type: str = 'ptw',
    entity_id: str = '',
    permit_id: str = '',
    actor_name: str = '',
) -> dict | None:
    clean_username = _username(username)
    clean_title = str(title or '').strip()
    if not clean_username or not clean_title:
        return None

    notifications = _load()
    record = {
        'id': str(uuid.uuid4()),
        'username': clean_username,
        'title': clean_title[:220],
        'message': str(message or '').strip()[:600],
        'link': str(link or '').strip()[:1000],
        'kind': str(kind or 'info').strip()[:24] or 'info',
        'entity_type': str(entity_type or '').strip()[:40],
        'entity_id': str(entity_id or '').strip()[:120],
        'permit_id': str(permit_id or '').strip()[:120],
        'actor_name': str(actor_name or '').strip()[:120],
        'created_at': _now(),
        'read_at': '',
    }
    notifications.append(record)
    notifications = sorted(notifications, key=_sort_key, reverse=True)[:2000]
    _save(notifications)
    return record


def create_notifications(
    usernames: list[str] | tuple[str, ...] | set[str],
    **kwargs,
) -> list[dict]:
    created: list[dict] = []
    seen: set[str] = set()
    for username in usernames or []:
        clean_username = _username(username)
        if not clean_username or clean_username in seen:
            continue
        seen.add(clean_username)
        notification = create_notification(clean_username, **kwargs)
        if notification:
            created.append(notification)
    return created


def list_notifications(
    username: str,
    *,
    unread_only: bool = False,
    limit: int = 20,
) -> list[dict]:
    clean_username = _username(username)
    if not clean_username:
        return []
    items = [
        _normalize_notification(notification, clean_username)
        for notification in _load()
        if _username(notification.get('username')) == clean_username
    ]
    if unread_only:
        items = [item for item in items if item.get('is_unread')]
    return items[: max(1, min(int(limit or 20), 100))]


def unread_count(username: str) -> int:
    clean_username = _username(username)
    if not clean_username:
        return 0
    return sum(
        1
        for notification in _load()
        if _username(notification.get('username')) == clean_username
        and not str(notification.get('read_at') or '').strip()
    )


def mark_notification_read(notification_id: str, username: str) -> dict | None:
    clean_username = _username(username)
    clean_id = str(notification_id or '').strip()
    if not clean_username or not clean_id:
        return None

    notifications = _load()
    updated = None
    for notification in notifications:
        if (
            str(notification.get('id') or '').strip() == clean_id
            and _username(notification.get('username')) == clean_username
        ):
            if not str(notification.get('read_at') or '').strip():
                notification['read_at'] = _now()
            updated = _normalize_notification(notification, clean_username)
            break
    if updated is None:
        return None
    _save(notifications)
    return updated


def mark_all_notifications_read(username: str) -> int:
    clean_username = _username(username)
    if not clean_username:
        return 0

    notifications = _load()
    changed = 0
    now = _now()
    for notification in notifications:
        if (
            _username(notification.get('username')) == clean_username
            and not str(notification.get('read_at') or '').strip()
        ):
            notification['read_at'] = now
            changed += 1
    if changed:
        _save(notifications)
    return changed
