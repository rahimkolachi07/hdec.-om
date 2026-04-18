"""
meeting_data.py  —  Data layer for the Meeting Room module.
All data is stored in meeting_data/ at the project root as JSON files.
"""
import json, uuid
from pathlib import Path
from datetime import datetime, timezone

BASE          = Path(__file__).resolve().parent.parent / 'meeting_data'
BASE.mkdir(exist_ok=True)

ROOMS_FILE           = BASE / 'rooms.json'
MESSAGES_FILE        = BASE / 'messages.json'
GROUPS_FILE          = BASE / 'groups.json'
FILES_FILE           = BASE / 'files.json'
CALLS_FILE           = BASE / 'calls.json'
PRESENCE_FILE        = BASE / 'presence.json'
ROOM_MEETINGS_FILE   = BASE / 'room_meetings.json'
ROOM_SIGNALS_FILE    = BASE / 'room_signals.json'
GLOBAL_PRESENCE_FILE = BASE / 'global_presence.json'
EMAIL_NOTIFIED_FILE  = BASE / 'email_notified.json'
FILES_DIR            = BASE / 'files'
FILES_DIR.mkdir(exist_ok=True)


def _now():
    return datetime.now(timezone.utc).isoformat()

def _uid():
    return str(uuid.uuid4())

def _load(path):
    if not path.exists():
        return []
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []

def _load_dict(path):
    if not path.exists():
        return {}
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def _save(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# ── PRESENCE ──────────────────────────────────────────────────────────────────

def update_presence(username: str):
    p = _load_dict(PRESENCE_FILE)
    p[username] = _now()
    _save(PRESENCE_FILE, p)

def get_online_users(threshold_seconds: int = 30) -> set:
    p = _load_dict(PRESENCE_FILE)
    now = datetime.now(timezone.utc)
    online = set()
    for u, ts in p.items():
        try:
            t = datetime.fromisoformat(ts)
            if (now - t).total_seconds() <= threshold_seconds:
                online.add(u)
        except Exception:
            pass
    return online


# ── ROOMS ─────────────────────────────────────────────────────────────────────

def get_rooms() -> list:
    return _load(ROOMS_FILE)

def get_room(room_id: str) -> dict | None:
    return next((r for r in get_rooms() if r.get('id') == room_id), None)

def create_room(data: dict) -> dict:
    rooms = get_rooms()
    host = data.get('host', '')
    room = {
        'id': _uid(),
        'name': data.get('name', 'New Room'),
        'title': data.get('title', ''),
        'host': host,
        'description': data.get('description', ''),
        'is_private': bool(data.get('is_private', False)),
        'password': data.get('password', ''),
        'participants': [host] if host else [],
        'created_at': _now(),
        'status': 'active',
    }
    rooms.append(room)
    _save(ROOMS_FILE, rooms)
    return room

def update_room(room_id: str, data: dict) -> dict | None:
    rooms = get_rooms()
    for r in rooms:
        if r.get('id') == room_id:
            for k in ('name', 'title', 'description', 'is_private', 'password', 'participants', 'status'):
                if k in data:
                    r[k] = data[k]
            _save(ROOMS_FILE, rooms)
            return r
    return None

def delete_room(room_id: str) -> bool:
    rooms = get_rooms()
    new = [r for r in rooms if r.get('id') != room_id]
    if len(new) == len(rooms):
        return False
    _save(ROOMS_FILE, new)
    return True

def join_room(room_id: str, username: str) -> dict | None:
    rooms = get_rooms()
    for r in rooms:
        if r.get('id') == room_id:
            if username not in r.get('participants', []):
                r.setdefault('participants', []).append(username)
            _save(ROOMS_FILE, rooms)
            return r
    return None

def leave_room(room_id: str, username: str) -> dict | None:
    rooms = get_rooms()
    for r in rooms:
        if r.get('id') == room_id:
            r['participants'] = [p for p in r.get('participants', []) if p != username]
            _save(ROOMS_FILE, rooms)
            return r
    return None


# ── MESSAGES ──────────────────────────────────────────────────────────────────

def get_messages(thread_id: str, limit: int = 100, after: str = '') -> list:
    msgs = [m for m in _load(MESSAGES_FILE) if m.get('thread_id') == thread_id]
    msgs.sort(key=lambda m: m.get('created_at', ''))
    if after:
        msgs = [m for m in msgs if m.get('created_at', '') > after]
    return msgs[-limit:]

def send_message(data: dict) -> dict:
    msgs = _load(MESSAGES_FILE)
    msg = {
        'id': _uid(),
        'thread_id': data.get('thread_id', ''),
        'sender': data.get('sender', ''),
        'content': data.get('content', ''),
        'type': data.get('type', 'text'),
        'file_id': data.get('file_id', ''),
        'created_at': _now(),
        'seen_by': [data.get('sender', '')],
        'translation_target': data.get('translation_target', ''),
        'translation_target_label': data.get('translation_target_label', ''),
        'translation_source_language': data.get('translation_source_language', ''),
        'translation_source_text': data.get('translation_source_text', ''),
        'translation_engine': data.get('translation_engine', ''),
    }
    msgs.append(msg)
    _save(MESSAGES_FILE, msgs)
    return msg

def mark_seen(thread_id: str, username: str):
    msgs = _load(MESSAGES_FILE)
    changed = False
    for m in msgs:
        if m.get('thread_id') == thread_id and username not in m.get('seen_by', []):
            m.setdefault('seen_by', []).append(username)
            changed = True
    if changed:
        _save(MESSAGES_FILE, msgs)

def get_unread_count(thread_id: str, username: str) -> int:
    return sum(
        1 for m in _load(MESSAGES_FILE)
        if m.get('thread_id') == thread_id
        and m.get('sender') != username
        and username not in m.get('seen_by', [])
    )

def thread_id_dm(user_a: str, user_b: str) -> str:
    return 'dm:' + ':'.join(sorted([user_a, user_b]))

def thread_id_group(group_id: str) -> str:
    return 'group:' + group_id

def thread_id_room(room_id: str) -> str:
    return 'room:' + room_id


# ── GROUPS ────────────────────────────────────────────────────────────────────

def get_groups() -> list:
    return _load(GROUPS_FILE)

def get_group(group_id: str) -> dict | None:
    return next((g for g in get_groups() if g.get('id') == group_id), None)

def create_group(data: dict) -> dict:
    groups = get_groups()
    me = data.get('created_by', '')
    members = list(data.get('members', []))
    if me and me not in members:
        members.insert(0, me)
    group = {
        'id': _uid(),
        'name': data.get('name', 'New Group'),
        'members': members,
        'admins': [me] if me else [],
        'created_by': me,
        'created_at': _now(),
    }
    groups.append(group)
    _save(GROUPS_FILE, groups)
    return group

def update_group(group_id: str, data: dict) -> dict | None:
    groups = get_groups()
    for g in groups:
        if g.get('id') == group_id:
            for k in ('name', 'members', 'admins'):
                if k in data:
                    g[k] = data[k]
            _save(GROUPS_FILE, groups)
            return g
    return None

def delete_group(group_id: str) -> bool:
    groups = get_groups()
    new = [g for g in groups if g.get('id') != group_id]
    if len(new) == len(groups):
        return False
    _save(GROUPS_FILE, new)
    return True


# ── FILES ─────────────────────────────────────────────────────────────────────

def get_files(thread_id: str = '') -> list:
    files = _load(FILES_FILE)
    if thread_id:
        files = [f for f in files if f.get('thread_id') == thread_id]
    return sorted(files, key=lambda f: f.get('created_at', ''), reverse=True)

def save_file_meta(data: dict) -> dict:
    files = _load(FILES_FILE)
    meta = {
        'id': data.get('id', _uid()),
        'name': data.get('name', ''),
        'size': data.get('size', 0),
        'mime': data.get('mime', 'application/octet-stream'),
        'sender': data.get('sender', ''),
        'thread_id': data.get('thread_id', ''),
        'path': data.get('path', ''),
        'created_at': _now(),
    }
    files.append(meta)
    _save(FILES_FILE, files)
    return meta

def get_file_meta(file_id: str) -> dict | None:
    return next((f for f in _load(FILES_FILE) if f.get('id') == file_id), None)


# ── CALLS  (WebRTC polling signaling) ─────────────────────────────────────────

def get_calls() -> list:
    return _load(CALLS_FILE)

def get_call(call_id: str) -> dict | None:
    return next((c for c in get_calls() if c.get('id') == call_id), None)

def create_call(data: dict) -> dict:
    calls = get_calls()
    call = {
        'id': _uid(),
        'caller': data.get('caller', ''),
        'callee': data.get('callee', ''),
        'call_type': data.get('call_type', 'audio'),
        'status': 'ringing',
        'sdp_offer': data.get('sdp_offer', ''),
        'sdp_answer': '',
        'caller_ice': [],
        'callee_ice': [],
        'created_at': _now(),
        'updated_at': _now(),
    }
    calls.append(call)
    _save(CALLS_FILE, calls)
    return call

def update_call(call_id: str, data: dict) -> dict | None:
    calls = get_calls()
    for c in calls:
        if c.get('id') == call_id:
            for k in ('status', 'sdp_answer'):
                if k in data:
                    c[k] = data[k]
            if 'caller_ice' in data and isinstance(data['caller_ice'], list):
                c.setdefault('caller_ice', []).extend(data['caller_ice'])
            if 'callee_ice' in data and isinstance(data['callee_ice'], list):
                c.setdefault('callee_ice', []).extend(data['callee_ice'])
            c['updated_at'] = _now()
            _save(CALLS_FILE, calls)
            return c
    return None

def get_pending_call_for(callee: str) -> dict | None:
    calls = get_calls()
    candidates = [c for c in calls if c.get('callee') == callee and c.get('status') == 'ringing']
    return candidates[-1] if candidates else None

def get_active_call_for(username: str) -> dict | None:
    calls = get_calls()
    active = [c for c in calls
              if c.get('status') == 'active'
              and (c.get('caller') == username or c.get('callee') == username)]
    return active[-1] if active else None

def cleanup_old_calls():
    """Remove calls older than 1 hour or in terminal states."""
    from datetime import timedelta
    calls = get_calls()
    cutoff = (datetime.now(timezone.utc) - timedelta(hours=1)).isoformat()
    calls = [c for c in calls if c.get('created_at', '') > cutoff or c.get('status') in ('ringing', 'active')]
    _save(CALLS_FILE, calls)


def get_room_meetings() -> list:
    meetings = _load(ROOM_MEETINGS_FILE)
    return sorted(meetings, key=lambda m: m.get('created_at', ''))


def get_room_meeting(meeting_id: str) -> dict | None:
    return next((m for m in get_room_meetings() if m.get('id') == meeting_id), None)


def get_live_room_meeting(room_id: str) -> dict | None:
    live = [m for m in get_room_meetings() if m.get('room_id') == room_id and m.get('status') == 'live']
    return live[-1] if live else None


def create_room_meeting(data: dict) -> dict:
    meetings = get_room_meetings()
    host = data.get('host', '')
    room_id = data.get('room_id', '')
    existing = next(
        (m for m in meetings if m.get('room_id') == room_id and m.get('status') == 'live'),
        None,
    )
    if existing:
        return existing
    meeting = {
        'id': _uid(),
        'room_id': room_id,
        'host': host,
        'title': data.get('title', ''),
        'call_type': data.get('call_type', 'video'),
        'participants': [host] if host else [],
        'status': 'live',
        'screen_sharing_by': '',
        'created_at': _now(),
        'updated_at': _now(),
    }
    meetings.append(meeting)
    _save(ROOM_MEETINGS_FILE, meetings)
    return meeting


def update_room_meeting(meeting_id: str, data: dict) -> dict | None:
    meetings = get_room_meetings()
    for meeting in meetings:
        if meeting.get('id') == meeting_id:
            for key in ('title', 'status', 'screen_sharing_by'):
                if key in data:
                    meeting[key] = data[key]
            if 'host' in data and data['host']:
                meeting['host'] = data['host']
            meeting['updated_at'] = _now()
            _save(ROOM_MEETINGS_FILE, meetings)
            return meeting
    return None


def join_room_meeting(meeting_id: str, username: str) -> dict | None:
    meetings = get_room_meetings()
    for meeting in meetings:
        if meeting.get('id') == meeting_id and meeting.get('status') == 'live':
            if username not in meeting.get('participants', []):
                meeting.setdefault('participants', []).append(username)
            meeting['updated_at'] = _now()
            _save(ROOM_MEETINGS_FILE, meetings)
            return meeting
    return None


def leave_room_meeting(meeting_id: str, username: str) -> dict | None:
    meetings = get_room_meetings()
    for meeting in meetings:
        if meeting.get('id') == meeting_id:
            participants = [p for p in meeting.get('participants', []) if p != username]
            meeting['participants'] = participants
            if meeting.get('screen_sharing_by') == username:
                meeting['screen_sharing_by'] = ''
            if participants:
                if meeting.get('host') == username:
                    meeting['host'] = participants[0]
            else:
                meeting['status'] = 'ended'
                meeting['screen_sharing_by'] = ''
            meeting['updated_at'] = _now()
            _save(ROOM_MEETINGS_FILE, meetings)
            return meeting
    return None


def get_room_signals(meeting_id: str = '', to_user: str = '', after: str = '') -> list:
    signals = _load(ROOM_SIGNALS_FILE)
    if meeting_id:
        signals = [s for s in signals if s.get('meeting_id') == meeting_id]
    if to_user:
        signals = [s for s in signals if s.get('to_user') == to_user]
    if after:
        signals = [s for s in signals if s.get('created_at', '') > after]
    signals.sort(key=lambda s: s.get('created_at', ''))
    return signals


def create_room_signal(data: dict) -> dict:
    signals = _load(ROOM_SIGNALS_FILE)
    signal = {
        'id': _uid(),
        'meeting_id': data.get('meeting_id', ''),
        'from_user': data.get('from_user', ''),
        'to_user': data.get('to_user', ''),
        'kind': data.get('kind', ''),
        'payload': data.get('payload', ''),
        'created_at': _now(),
    }
    signals.append(signal)
    _save(ROOM_SIGNALS_FILE, signals)
    return signal


def cleanup_old_room_meetings():
    from datetime import timedelta

    cutoff = (datetime.now(timezone.utc) - timedelta(hours=8)).isoformat()

    meetings = [
        m for m in get_room_meetings()
        if m.get('status') == 'live' or m.get('updated_at', m.get('created_at', '')) > cutoff
    ]
    live_ids = {m.get('id') for m in meetings if m.get('status') == 'live'}
    _save(ROOM_MEETINGS_FILE, meetings)

    signals = [
        s for s in _load(ROOM_SIGNALS_FILE)
        if s.get('meeting_id') in live_ids or s.get('created_at', '') > cutoff
    ]
    _save(ROOM_SIGNALS_FILE, signals)


# ── GLOBAL PRESENCE (all pages) ───────────────────────────────────────────────

def update_global_presence(username: str):
    p = _load_dict(GLOBAL_PRESENCE_FILE)
    p[username] = _now()
    _save(GLOBAL_PRESENCE_FILE, p)

def get_globally_online_users(threshold_seconds: int = 90) -> set:
    p = _load_dict(GLOBAL_PRESENCE_FILE)
    now = datetime.now(timezone.utc)
    online = set()
    for u, ts in p.items():
        try:
            t = datetime.fromisoformat(ts)
            if (now - t).total_seconds() <= threshold_seconds:
                online.add(u)
        except Exception:
            pass
    return online

def is_globally_online(username: str, threshold_seconds: int = 90) -> bool:
    p = _load_dict(GLOBAL_PRESENCE_FILE)
    ts = p.get(username)
    if not ts:
        return False
    try:
        t = datetime.fromisoformat(ts)
        return (datetime.now(timezone.utc) - t).total_seconds() <= threshold_seconds
    except Exception:
        return False


# ── EMAIL NOTIFICATION DEDUP ──────────────────────────────────────────────────

def should_send_message_email(username: str, thread_id: str, cooldown_minutes: int = 15) -> bool:
    notified = _load_dict(EMAIL_NOTIFIED_FILE)
    last_ts = notified.get(username, {}).get(thread_id)
    if not last_ts:
        return True
    try:
        t = datetime.fromisoformat(last_ts)
        return (datetime.now(timezone.utc) - t).total_seconds() / 60 >= cooldown_minutes
    except Exception:
        return True

def record_message_email_sent(username: str, thread_id: str):
    notified = _load_dict(EMAIL_NOTIFIED_FILE)
    notified.setdefault(username, {})[thread_id] = _now()
    _save(EMAIL_NOTIFIED_FILE, notified)


# ── UNREAD SUMMARY ────────────────────────────────────────────────────────────

def get_unread_threads_for(username: str) -> list:
    """Return threads with unread messages and latest preview for a user."""
    result = []
    try:
        from .auth_utils import get_all_users as _get_all_users
        all_users = _get_all_users()
    except Exception:
        all_users = []

    for u in all_users:
        if u['username'] == username:
            continue
        tid = thread_id_dm(username, u['username'])
        count = get_unread_count(tid, username)
        if count:
            last = get_messages(tid, limit=1)
            result.append({
                'thread_id': tid,
                'type': 'dm',
                'peer_name': u.get('name') or u['username'],
                'count': count,
                'preview': last[-1].get('content', '')[:120] if last else '',
                'sender': last[-1].get('sender', '') if last else '',
            })

    for g in get_groups():
        if username not in g.get('members', []):
            continue
        tid = thread_id_group(g['id'])
        count = get_unread_count(tid, username)
        if count:
            last = get_messages(tid, limit=1)
            result.append({
                'thread_id': tid,
                'type': 'group',
                'group_name': g.get('name', 'Group'),
                'count': count,
                'preview': last[-1].get('content', '')[:120] if last else '',
                'sender': last[-1].get('sender', '') if last else '',
            })
    return result
