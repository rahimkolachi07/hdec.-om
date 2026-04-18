"""
meeting_views.py  —  Views and API endpoints for the Meeting Room module.
Isolated to Meeting Room only. No shared code is modified.
"""
import json, uuid as _uuid_mod
from pathlib import Path
from functools import wraps

from django.shortcuts import render, redirect
from django.http import JsonResponse, FileResponse, Http404
from django.views.decorators.csrf import csrf_exempt

from .auth_utils import get_all_users, normalize_user_state
from .meeting_data import (
    update_presence, get_online_users,
    get_rooms, get_room, create_room, update_room, delete_room, join_room, leave_room,
    get_messages, send_message, mark_seen, get_unread_count,
    thread_id_dm, thread_id_group, thread_id_room,
    get_groups, get_group, create_group, update_group, delete_group,
    get_files, get_file_meta, save_file_meta, FILES_DIR,
    get_calls, get_call, create_call, update_call,
    get_pending_call_for, get_active_call_for, cleanup_old_calls,
    get_room_meetings, get_room_meeting, get_live_room_meeting, create_room_meeting,
    update_room_meeting, join_room_meeting, leave_room_meeting,
    get_room_signals, create_room_signal, cleanup_old_room_meetings,
    update_global_presence, is_globally_online, get_globally_online_users,
    should_send_message_email, record_message_email_sent,
    get_unread_threads_for,
)
from .translation_utils import (
    TRANSLATION_LANGUAGES,
    TranslationError,
    extract_translation_directive,
    translate_text,
)
from .openai_realtime import OpenAIRealtimeError, create_call_translation_session


def _get_user(request):
    return normalize_user_state(request.session.get('hdec_user'))


def _login_required(fn):
    @wraps(fn)
    def wrapper(request, *args, **kwargs):
        if not _get_user(request):
            return redirect('/login/')
        return fn(request, *args, **kwargs)
    return wrapper


def _err(msg, status=400):
    return JsonResponse({'error': msg}, status=status)


def _user_list_with_status(me: str) -> list:
    online = get_online_users()
    return [
        {
            'username': u['username'],
            'name': u['name'],
            'role': u.get('role', ''),
            'online': u['username'] in online,
        }
        for u in get_all_users()
        if u['username'] != me
    ]


def _hydrate_messages(messages: list[dict]) -> list[dict]:
    hydrated = []
    for msg in messages:
        item = dict(msg)
        if item.get('type') == 'file' and item.get('file_id'):
            meta = get_file_meta(item['file_id'])
            if meta:
                item['mime'] = meta.get('mime', 'application/octet-stream')
                item['file_size'] = meta.get('size', 0)
                item['file_name'] = meta.get('name', item.get('content', ''))
                item['file_url'] = f"/api/meeting/files/{item['file_id']}/"
                item['file_inline_url'] = f"/api/meeting/files/{item['file_id']}/?inline=1"
        hydrated.append(item)
    return hydrated


def _live_room_meetings() -> list:
    return [m for m in get_room_meetings() if m.get('status') == 'live']


def _room_access(room_id: str, username: str) -> tuple[dict | None, str | None]:
    room = get_room(room_id)
    if not room:
        return None, 'Room not found'
    if username not in room.get('participants', []):
        return None, 'Join the room first'
    return room, None


# ── MAIN PAGE ─────────────────────────────────────────────────────────────────

@_login_required
def meeting_hub(request, country_id='', project_id=''):
    user = _get_user(request)
    me = user.get('username', '') if user else ''
    if me:
        update_presence(me)
        cleanup_old_calls()
        cleanup_old_room_meetings()

    all_users_json = json.dumps(_user_list_with_status(me))
    groups = [g for g in get_groups() if me in g.get('members', [])]
    rooms  = [r for r in get_rooms()  if r.get('status') == 'active']

    # unread counts per thread
    unread = {}
    for u in get_all_users():
        if u['username'] == me:
            continue
        tid = thread_id_dm(me, u['username'])
        c = get_unread_count(tid, me)
        if c:
            unread[tid] = c
    for g in groups:
        tid = thread_id_group(g['id'])
        c = get_unread_count(tid, me)
        if c:
            unread[tid] = c

    return render(request, 'core/meeting_hub.html', {
        'user': user,
        'me': me,
        'me_name': user.get('name', me) if user else me,
        'all_users_json': all_users_json,
        'groups_json': json.dumps(groups),
        'rooms_json': json.dumps(rooms),
        'room_meetings_json': json.dumps(_live_room_meetings()),
        'translation_languages_json': json.dumps(
            [{'code': item['code'], 'label': item['label']} for item in TRANSLATION_LANGUAGES]
        ),
        'unread_json': json.dumps(unread),
        'country_id': country_id,
        'project_id': project_id,
        'csrf_token': request.META.get('CSRF_COOKIE', ''),
    })


# ── PRESENCE ──────────────────────────────────────────────────────────────────

@csrf_exempt
def meeting_api_presence(request):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')
    update_presence(me)
    online = list(get_online_users())
    return JsonResponse({'online': online})


# ── USERS ─────────────────────────────────────────────────────────────────────

@csrf_exempt
def meeting_api_users(request):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')
    return JsonResponse({'users': _user_list_with_status(me)})


# ── ROOMS ─────────────────────────────────────────────────────────────────────

@csrf_exempt
def meeting_api_rooms(request):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')

    if request.method == 'GET':
        return JsonResponse({'rooms': [r for r in get_rooms() if r.get('status') == 'active']})

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
        except Exception:
            return _err('Invalid JSON')
        data['host'] = me
        room = create_room(data)
        return JsonResponse({'room': room}, status=201)

    return _err('Method not allowed', 405)


@csrf_exempt
def meeting_api_room(request, room_id):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')

    if request.method == 'DELETE':
        room = get_room(room_id)
        if not room:
            return _err('Not found', 404)
        if room.get('host') != me and user.get('role') != 'admin':
            return _err('Permission denied', 403)
        delete_room(room_id)
        return JsonResponse({'ok': True})

    if request.method == 'PATCH':
        try:
            data = json.loads(request.body)
        except Exception:
            return _err('Invalid JSON')
        room = update_room(room_id, data)
        if not room:
            return _err('Not found', 404)
        return JsonResponse({'room': room})

    return _err('Method not allowed', 405)


@csrf_exempt
def meeting_api_room_join(request, room_id):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    room = join_room(room_id, user.get('username', ''))
    return JsonResponse({'room': room}) if room else _err('Not found', 404)


@csrf_exempt
def meeting_api_room_leave(request, room_id):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    room = leave_room(room_id, user.get('username', ''))
    return JsonResponse({'room': room}) if room else _err('Not found', 404)


# ── MESSAGES ──────────────────────────────────────────────────────────────────

@csrf_exempt
def meeting_api_messages(request):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')

    if request.method == 'GET':
        thread_id = request.GET.get('thread', '')
        after     = request.GET.get('after', '')
        if not thread_id:
            return _err('thread required')
        mark_seen(thread_id, me)
        msgs = _hydrate_messages(get_messages(thread_id, limit=100, after=after))
        return JsonResponse({'messages': msgs})

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
        except Exception:
            return _err('Invalid JSON')
        data['sender'] = me
        if data.get('type', 'text') == 'text':
            content = str(data.get('content', '')).strip()
            target_language = str(data.get('translate_to', '')).strip()
            if not target_language:
                inline_target, content = extract_translation_directive(content)
                if inline_target:
                    target_language = inline_target['code']
            data['content'] = content
            if target_language:
                try:
                    translated = translate_text(content, target_language)
                except TranslationError as exc:
                    return _err(str(exc), 503)
                data['content'] = translated['text']
                data['translation_target'] = translated['target_language']
                data['translation_target_label'] = translated['target_label']
                data['translation_source_language'] = translated['source_language']
                data['translation_source_text'] = content
                data['translation_engine'] = translated['engine']
        msg = send_message(data)
        _notify_message_recipients_by_email(me, data.get('thread_id', ''), msg)
        return JsonResponse({'message': msg}, status=201)

    return _err('Method not allowed', 405)


# ── GROUPS ────────────────────────────────────────────────────────────────────

@csrf_exempt
def meeting_api_groups(request):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')

    if request.method == 'GET':
        groups = [g for g in get_groups() if me in g.get('members', [])]
        return JsonResponse({'groups': groups})

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
        except Exception:
            return _err('Invalid JSON')
        data['created_by'] = me
        group = create_group(data)
        return JsonResponse({'group': group}, status=201)

    return _err('Method not allowed', 405)


@csrf_exempt
def meeting_api_group(request, group_id):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')

    if request.method == 'DELETE':
        group = get_group(group_id)
        if not group:
            return _err('Not found', 404)
        if group.get('created_by') != me and user.get('role') != 'admin':
            return _err('Permission denied', 403)
        delete_group(group_id)
        return JsonResponse({'ok': True})

    if request.method == 'PATCH':
        try:
            data = json.loads(request.body)
        except Exception:
            return _err('Invalid JSON')
        group = update_group(group_id, data)
        if not group:
            return _err('Not found', 404)
        return JsonResponse({'group': group})

    return _err('Method not allowed', 405)


# ── FILES ─────────────────────────────────────────────────────────────────────

@csrf_exempt
def meeting_api_files(request):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')

    if request.method == 'GET':
        thread_id = request.GET.get('thread', '')
        return JsonResponse({'files': get_files(thread_id)})

    if request.method == 'POST':
        f = request.FILES.get('file')
        thread_id = request.POST.get('thread_id', '')
        if not f:
            return _err('No file')
        fid = str(_uuid_mod.uuid4())
        dest_dir = FILES_DIR / fid
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest_file = dest_dir / f.name
        with open(dest_file, 'wb') as out:
            for chunk in f.chunks():
                out.write(chunk)
        meta = save_file_meta({
            'id': fid,
            'name': f.name,
            'size': f.size,
            'mime': f.content_type or 'application/octet-stream',
            'sender': me,
            'thread_id': thread_id,
            'path': str(dest_file),
        })
        if thread_id:
            send_message({
                'thread_id': thread_id,
                'sender': me,
                'content': f.name,
                'type': 'file',
                'file_id': fid,
            })
        return JsonResponse({'file': meta}, status=201)

    return _err('Method not allowed', 405)


@csrf_exempt
def meeting_api_file_download(request, file_id):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    meta = get_file_meta(file_id)
    if not meta:
        raise Http404
    path = Path(meta['path'])
    if not path.exists():
        raise Http404
    mime = meta.get('mime', 'application/octet-stream')
    resp = FileResponse(open(path, 'rb'), content_type=mime)
    inline = request.GET.get('inline', '').lower() in ('1', 'true', 'yes')
    disp = 'inline' if inline else 'attachment'
    resp['Content-Disposition'] = f'{disp}; filename="{meta["name"]}"'
    if inline:
        resp['X-Frame-Options'] = 'SAMEORIGIN'
    return resp


# ── CALLS  (WebRTC polling signaling) ─────────────────────────────────────────

@csrf_exempt
def meeting_api_calls(request):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')

    if request.method == 'GET':
        pending = get_pending_call_for(me)
        active  = get_active_call_for(me)
        return JsonResponse({'pending': pending, 'active': active})

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
        except Exception:
            return _err('Invalid JSON')
        data['caller'] = me
        call = create_call(data)
        _notify_call_by_email(me, data.get('callee', ''), data.get('call_type', 'audio'))
        return JsonResponse({'call': call}, status=201)

    return _err('Method not allowed', 405)


@csrf_exempt
def meeting_api_realtime_session(request):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)

    if request.method != 'POST':
        return _err('Method not allowed', 405)

    try:
        data = json.loads(request.body)
    except Exception:
        return _err('Invalid JSON')

    sdp = str(data.get('sdp', '')).strip()
    target_language = str(data.get('target_language', '')).strip()
    if not sdp:
        return _err('sdp required')
    if not target_language:
        return _err('target_language required')

    try:
        answer_sdp, target = create_call_translation_session(sdp, target_language)
    except OpenAIRealtimeError as exc:
        message = str(exc)
        status = 503
        if 'Unsupported translation language' in message:
            status = 400
        elif 'not configured' in message:
            status = 500
        return _err(message, status)

    return JsonResponse({
        'sdp': answer_sdp,
        'target_language': target['code'],
        'target_label': target['label'],
    })


@csrf_exempt
def meeting_api_call(request, call_id):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)

    if request.method == 'GET':
        call = get_call(call_id)
        return JsonResponse({'call': call}) if call else _err('Not found', 404)

    if request.method == 'PATCH':
        try:
            data = json.loads(request.body)
        except Exception:
            return _err('Invalid JSON')
        call = update_call(call_id, data)
        return JsonResponse({'call': call}) if call else _err('Not found', 404)

    return _err('Method not allowed', 405)


@csrf_exempt
def meeting_api_room_meetings(request):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')

    if request.method == 'GET':
        room_id = request.GET.get('room_id', '')
        if room_id:
            room, err = _room_access(room_id, me)
            if not room:
                return _err(err or 'Room not found', 403 if err == 'Join the room first' else 404)
            return JsonResponse({'meeting': get_live_room_meeting(room_id), 'room': room})
        return JsonResponse({'meetings': _live_room_meetings()})

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
        except Exception:
            return _err('Invalid JSON')
        room_id = data.get('room_id', '')
        room, err = _room_access(room_id, me)
        if not room:
            return _err(err or 'Room not found', 403 if err == 'Join the room first' else 404)
        meeting = create_room_meeting({
            'room_id': room_id,
            'host': me,
            'title': data.get('title') or room.get('title') or room.get('name', ''),
            'call_type': data.get('call_type', 'video'),
        })
        if me not in meeting.get('participants', []):
            meeting = join_room_meeting(meeting['id'], me) or meeting
        return JsonResponse({'meeting': meeting}, status=201)

    return _err('Method not allowed', 405)


@csrf_exempt
def meeting_api_room_meeting(request, meeting_id):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')

    meeting = get_room_meeting(meeting_id)
    if not meeting:
        return _err('Not found', 404)

    room, err = _room_access(meeting.get('room_id', ''), me)
    if not room:
        return _err(err or 'Room not found', 403 if err == 'Join the room first' else 404)

    if request.method == 'GET':
        return JsonResponse({'meeting': meeting, 'room': room})

    if request.method == 'PATCH':
        try:
            data = json.loads(request.body)
        except Exception:
            return _err('Invalid JSON')

        action = data.get('action', '')
        if action == 'join':
            meeting = join_room_meeting(meeting_id, me)
        elif action == 'leave':
            meeting = leave_room_meeting(meeting_id, me)
        elif action == 'share':
            if me not in meeting.get('participants', []):
                return _err('Join the meeting first', 403)
            share_on = bool(data.get('enabled'))
            meeting = update_room_meeting(meeting_id, {
                'screen_sharing_by': me if share_on else '',
            })
        elif action == 'end':
            if me != meeting.get('host') and user.get('role') != 'admin':
                return _err('Permission denied', 403)
            meeting = update_room_meeting(meeting_id, {
                'status': 'ended',
                'screen_sharing_by': '',
            })
        else:
            meeting = update_room_meeting(meeting_id, data)

        return JsonResponse({'meeting': meeting}) if meeting else _err('Not found', 404)

    return _err('Method not allowed', 405)


@csrf_exempt
def meeting_api_room_signals(request):
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')

    if request.method == 'GET':
        meeting_id = request.GET.get('meeting_id', '')
        after = request.GET.get('after', '')
        if not meeting_id:
            return _err('meeting_id required')
        meeting = get_room_meeting(meeting_id)
        if not meeting:
            return _err('Not found', 404)
        room, err = _room_access(meeting.get('room_id', ''), me)
        if not room:
            return _err(err or 'Room not found', 403 if err == 'Join the room first' else 404)
        return JsonResponse({'signals': get_room_signals(meeting_id=meeting_id, to_user=me, after=after)})

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
        except Exception:
            return _err('Invalid JSON')
        meeting_id = data.get('meeting_id', '')
        meeting = get_room_meeting(meeting_id)
        if not meeting:
            return _err('Not found', 404)
        room, err = _room_access(meeting.get('room_id', ''), me)
        if not room:
            return _err(err or 'Room not found', 403 if err == 'Join the room first' else 404)
        if me not in meeting.get('participants', []):
            return _err('Join the meeting first', 403)
        signal = create_room_signal({
            'meeting_id': meeting_id,
            'from_user': me,
            'to_user': data.get('to_user', ''),
            'kind': data.get('kind', ''),
            'payload': data.get('payload', ''),
        })
        return JsonResponse({'signal': signal}, status=201)

    return _err('Method not allowed', 405)


# ── POLL  (combined polling endpoint for efficiency) ──────────────────────────

@csrf_exempt
def meeting_api_poll(request):
    """Single polling endpoint: returns messages, online users, pending calls."""
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')
    update_presence(me)
    cleanup_old_room_meetings()

    thread_id = request.GET.get('thread', '')
    after     = request.GET.get('after', '')

    new_msgs = []
    if thread_id:
        mark_seen(thread_id, me)
        new_msgs = _hydrate_messages(get_messages(thread_id, limit=50, after=after))

    # unread counts for sidebar badge
    all_users = get_all_users()
    groups    = [g for g in get_groups() if me in g.get('members', [])]
    unread    = {}
    for u in all_users:
        if u['username'] == me:
            continue
        tid = thread_id_dm(me, u['username'])
        c = get_unread_count(tid, me)
        if c:
            unread[tid] = c
    for g in groups:
        tid = thread_id_group(g['id'])
        c = get_unread_count(tid, me)
        if c:
            unread[tid] = c

    online  = list(get_online_users())
    pending = get_pending_call_for(me)
    active  = get_active_call_for(me)

    return JsonResponse({
        'messages': new_msgs,
        'online':   online,
        'unread':   unread,
        'pending_call': pending,
        'active_call':  active,
        'groups':   groups,
        'rooms':    [r for r in get_rooms() if r.get('status') == 'active'],
        'room_meetings': _live_room_meetings(),
    })


# ── EMAIL HELPERS ─────────────────────────────────────────────────────────────

def _notify_message_recipients_by_email(sender: str, thread_id: str, msg: dict):
    """Create bell notifications for all recipients; email offline ones."""
    try:
        from .email_utils import notify_meeting_message
        from .notification_utils import create_notification
        meeting_online = get_online_users(threshold_seconds=10)  # only skip bell if actively in meeting room right now

        all_users = get_all_users()
        user_map = {u['username']: u for u in all_users}
        sender_name = user_map.get(sender, {}).get('name') or sender
        content_preview = str(msg.get('content', ''))[:150]

        recipients = []
        thread_name = sender_name

        if thread_id.startswith('dm:'):
            parts = thread_id[3:].split(':')
            recipients = [p for p in parts if p != sender]
        elif thread_id.startswith('group:'):
            group = get_group(thread_id[6:])
            if group:
                recipients = [m for m in group.get('members', []) if m != sender]
                thread_name = group.get('name', 'Group')

        for recipient in recipients:
            # Always create a bell notification unless they're actively in the meeting room
            if recipient not in meeting_online:
                create_notification(
                    recipient,
                    title=f'💬 {sender_name}',
                    message=content_preview,
                    link='/meeting-room/',
                    kind='info',
                    entity_type='meeting',
                    actor_name=sender_name,
                )

            # Email only if completely offline (not on any page)
            if is_globally_online(recipient):
                continue
            if not should_send_message_email(recipient, thread_id):
                continue
            email = user_map.get(recipient, {}).get('email', '')
            if not email:
                continue
            notify_meeting_message(email, sender_name, thread_name, content_preview)
            record_message_email_sent(recipient, thread_id)
    except Exception:
        pass


def _notify_call_by_email(caller: str, callee: str, call_type: str):
    """Send email to callee if they are offline when a call is initiated."""
    try:
        if is_globally_online(callee):
            return
        from .email_utils import notify_meeting_call
        all_users = get_all_users()
        user_map = {u['username']: u for u in all_users}
        caller_name = user_map.get(caller, {}).get('name') or caller
        email = user_map.get(callee, {}).get('email', '')
        if email:
            notify_meeting_call(email, caller_name, call_type)
    except Exception:
        pass


# ── GLOBAL ALERTS (cross-module polling) ──────────────────────────────────────

@csrf_exempt
def meeting_api_global_alerts(request):
    """Lightweight endpoint polled by every page to surface meeting notifications."""
    user = _get_user(request)
    if not user:
        return _err('Unauthorized', 401)
    me = user.get('username', '')

    update_global_presence(me)

    all_users = get_all_users()
    user_map = {u['username']: u for u in all_users}

    # Online users (across all modules via global presence)
    globally_online = get_globally_online_users()
    online_users = [
        {'username': u, 'name': user_map.get(u, {}).get('name') or u}
        for u in sorted(globally_online)
        if u != me
    ]

    pending = get_pending_call_for(me)
    unread_threads = get_unread_threads_for(me)
    unread_total = sum(t['count'] for t in unread_threads)

    call_data = None
    if pending:
        caller_user = user_map.get(pending.get('caller', ''), {})
        call_data = dict(pending)
        call_data['caller_name'] = caller_user.get('name') or pending.get('caller', '')

    return JsonResponse({
        'pending_call':   call_data,
        'unread_total':   unread_total,
        'unread_threads': unread_threads,
        'online_users':   online_users,
        'online_count':   len(globally_online),
    })
