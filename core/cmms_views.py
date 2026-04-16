"""
CMMS Views — monthly schedule, activity work page, and API endpoints.
"""
import json
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import urlencode, urlparse
from django.shortcuts import render, redirect
from django.http import JsonResponse, Http404, HttpResponse
from django.views.decorators.csrf import csrf_exempt
from .auth_utils import (
    get_user_detail,
    get_users_by_role,
    has_permission,
    normalize_user_state,
    can_access_country,
    can_access_project,
)
from .project_utils import get_countries
from .email_utils import (
    notify_permit_closed,
    notify_permit_closure_hse_required,
    notify_permit_closure_requested,
    notify_permit_created,
    notify_permit_issued,
    notify_permit_ready_to_proceed,
)
from .notification_utils import (
    create_notifications,
    list_notifications,
    mark_all_notifications_read,
    mark_notification_read,
    unread_count,
)

from .cmms_utils import (
    get_activity, get_activities_for_month,
    activity_occurs_on_date,
    create_activity, update_activity, delete_activity, delete_activity_occurrence,
    get_record, get_record_for_activity_date, start_record, update_record,
    save_photo, delete_photo,
    parse_excel_checklist, generate_zip,
    get_checklist_files, resolve_checklist_path, ensure_activity_checklist, save_checklist_file,
    get_activity_sheet_label, get_activity_sheet_link, get_all_checklist_activities, get_activity_permit_options,
)
from .cmms_ptw_utils import (
    annotate_permit,
    application_is_active,
    build_permit_docx,
    can_close_hse,
    can_delete_permit,
    can_close_issuer,
    can_close_receiver,
    can_edit_application,
    can_hse_approve,
    can_issue_permit,
    can_receiver_unlock,
    create_or_get_record_permit,
    delete_permit,
    ensure_final_permit_pdf,
    get_permit,
    get_permit_for_record,
    is_cmms_permit,
    list_permits,
    permit_filename,
    update_permit,
)


# ── Auth helpers ──────────────────────────────────────────────────────────────
_MAINTENANCE_PATH_RE = re.compile(r'^/p/([^/]+)/([^/]+)/maintenance/?$')


def _get_user(request):
    user = normalize_user_state(request.session.get('hdec_user'))
    if user != request.session.get('hdec_user'):
        request.session['hdec_user'] = user
        request.session.modified = True
    return user


def _is_api_request(request) -> bool:
    return bool(
        request.path.startswith('/api/')
        or request.method != 'GET'
        or request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        or 'application/json' in request.headers.get('Accept', '')
    )


def _has_any_project_access(user: dict | None) -> bool:
    if not user:
        return False
    if user.get('role') == 'admin':
        return True
    for country in get_countries():
        country_id = country.get('id', '')
        if not can_access_country(user, country_id):
            continue
        for project in country.get('projects', []):
            if can_access_project(user, country_id, project.get('id', '')):
                return True
    return False


def _require_login(request):
    if not _get_user(request):
        if _is_api_request(request):
            return JsonResponse({'error': 'Login required'}, status=401)
        return redirect('/login/')
    return None


def _require_module_access(request, module_id: str, level: str = 'view'):
    redir = _require_login(request)
    if redir:
        return redir
    user = _get_user(request) or {}
    if _has_any_project_access(user) and has_permission(user, module_id, level):
        return None
    if _is_api_request(request):
        return JsonResponse({'error': 'Forbidden'}, status=403)
    return redirect('/')


def _maintenance_url(country_id: str, project_id: str) -> str:
    return f'/p/{country_id}/{project_id}/maintenance/'


def _accessible_maintenance_url(user: dict | None, raw_value: str) -> str:
    if not raw_value:
        return ''
    try:
        path = urlparse(raw_value).path or ''
    except Exception:
        path = str(raw_value or '')
    match = _MAINTENANCE_PATH_RE.match(path)
    if not match:
        return ''
    country_id, project_id = match.groups()
    if user and user.get('role') == 'admin':
        return _maintenance_url(country_id, project_id)
    if user and can_access_country(user, country_id) and can_access_project(user, country_id, project_id):
        return _maintenance_url(country_id, project_id)
    return ''


def _default_maintenance_url(user: dict | None) -> str:
    for country in get_countries():
        country_id = country.get('id', '')
        if user and user.get('role') != 'admin' and not can_access_country(user, country_id):
            continue
        for project in country.get('projects', []):
            project_id = project.get('id', '')
            if user and user.get('role') != 'admin' and not can_access_project(user, country_id, project_id):
                continue
            return _maintenance_url(country_id, project_id)
    return '/'


def _default_project_cmms_url(user: dict | None, suffix: str = '') -> str:
    clean_suffix = str(suffix or '').lstrip('/')
    for country in get_countries():
        country_id = country.get('id', '')
        if user and user.get('role') != 'admin' and not can_access_country(user, country_id):
            continue
        for project in country.get('projects', []):
            project_id = project.get('id', '')
            if user and user.get('role') != 'admin' and not can_access_project(user, country_id, project_id):
                continue
            base = f'/p/{country_id}/{project_id}/maintenance/cmms/'
            return f'{base}{clean_suffix}' if clean_suffix else base
    return '/cmms/'


def _cmms_back_url(request, user: dict | None) -> str:
    for raw_value in (
        request.GET.get('back', ''),
        request.META.get('HTTP_REFERER', ''),
        request.session.get('cmms_back_url', ''),
    ):
        resolved = _accessible_maintenance_url(user, raw_value)
        if resolved:
            if request.session.get('cmms_back_url') != resolved:
                request.session['cmms_back_url'] = resolved
                request.session.modified = True
            return resolved
    fallback = _default_maintenance_url(user)
    if request.session.get('cmms_back_url') != fallback:
        request.session['cmms_back_url'] = fallback
        request.session.modified = True
    return fallback


def _ctx(request, extra=None):
    user = _get_user(request)
    ctx = {
        'user': user,
        'csrf_token': request.META.get('CSRF_COOKIE', ''),
        'cmms_back_url': _cmms_back_url(request, user),
    }
    if extra:
        ctx.update(extra)
    return ctx


def _user_email(username: str) -> str:
    if not username:
        return ''
    user = get_user_detail(username)
    return (user or {}).get('email', '')


def _dedupe_users(users: list) -> list:
    seen = set()
    result = []
    for user in users:
        username = str((user or {}).get('username', '')).strip().lower()
        email = str((user or {}).get('email', '')).strip().lower()
        key = username or email
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(user)
    return result


def _dedupe_usernames(usernames: list[str]) -> list[str]:
    seen = set()
    result = []
    for username in usernames:
        clean = str(username or '').strip().lower()
        if not clean or clean in seen:
            continue
        seen.add(clean)
        result.append(clean)
    return result


def _issuer_users() -> list:
    users = get_users_by_role('operation_engineer')
    if users:
        return _dedupe_users(users)
    return _dedupe_users(get_users_by_role('admin'))


def _hse_users() -> list:
    users = get_users_by_role('hse_engineer')
    if users:
        return _dedupe_users(users)
    return _dedupe_users(get_users_by_role('admin'))


def _issuer_notification_users() -> list:
    return [user for user in _issuer_users() if user.get('email')]


def _hse_notification_users() -> list:
    return [user for user in _hse_users() if user.get('email')]


def _usernames_for_users(users: list) -> list[str]:
    return _dedupe_usernames([
        str((user or {}).get('username', '')).strip().lower()
        for user in users or []
    ])


def _receiver_and_maintenance_emails(receiver_username: str) -> list:
    emails = []
    receiver_email = _user_email(receiver_username)
    if receiver_email:
        emails.append(receiver_email)
    emails.extend(
        u.get('email', '')
        for u in get_users_by_role('maintenance_engineer')
        if u.get('email')
    )
    return list(dict.fromkeys(
        str(email).strip()
        for email in emails
        if email and '@' in str(email)
    ))


def _receiver_and_maintenance_usernames(receiver_username: str) -> list[str]:
    usernames = [receiver_username]
    usernames.extend(
        str((user or {}).get('username', '')).strip().lower()
        for user in get_users_by_role('maintenance_engineer')
    )
    return _dedupe_usernames(usernames)


def _closure_notification_emails(permit: dict | None) -> list:
    permit = permit or {}
    emails = _receiver_and_maintenance_emails(permit.get('receiver_username', ''))
    emails.extend([
        _user_email(permit.get('issuer_username', '')),
        _user_email(permit.get('hse_username', '')),
    ])
    return list(dict.fromkeys(
        str(email).strip()
        for email in emails
        if email and '@' in str(email)
    ))


def _closure_notification_usernames(permit: dict | None) -> list[str]:
    permit = permit or {}
    usernames = _receiver_and_maintenance_usernames(permit.get('receiver_username', ''))
    usernames.extend([
        permit.get('issuer_username', ''),
        permit.get('hse_username', ''),
    ])
    return _dedupe_usernames(usernames)


def _ptw_link(permit_id: str, *, focus: str = '', decision: str = '') -> str:
    clean_id = str(permit_id or '').strip()
    if not clean_id:
        return '/cmms/ptw/'
    params = {}
    if focus:
        params['focus'] = focus
    if decision:
        params['decision'] = decision
    query = urlencode(params)
    base = f'/cmms/ptw/{clean_id}/'
    return f'{base}?{query}' if query else base


def _push_ptw_notifications(
    usernames: list[str],
    *,
    title: str,
    message: str,
    permit: dict | None,
    focus: str = '',
    kind: str = 'info',
    actor_name: str = '',
) -> None:
    permit = permit or {}
    create_notifications(
        usernames,
        title=title,
        message=message,
        link=_ptw_link(permit.get('id', ''), focus=focus),
        kind=kind,
        entity_type='ptw',
        entity_id=permit.get('id', ''),
        permit_id=permit.get('id', ''),
        actor_name=actor_name,
    )


def notifications_api(request):
    redir = _require_login(request)
    if redir:
        return redir
    user = _get_user(request) or {}
    username = str(user.get('username', '')).strip().lower()
    if not username:
        return JsonResponse({'error': 'Login required'}, status=401)

    if request.method == 'GET':
        try:
            limit = max(1, min(int(request.GET.get('limit', '20')), 100))
        except Exception:
            limit = 20
        return JsonResponse({
            'notifications': list_notifications(username, limit=limit),
            'unread_count': unread_count(username),
        })

    if request.method != 'POST':
        return JsonResponse({'error': 'Method not allowed'}, status=405)

    try:
        data = json.loads(request.body or '{}')
    except Exception:
        return JsonResponse({'error': 'Invalid JSON'}, status=400)

    action = str(data.get('action', '') or '').strip()
    if action == 'read':
        notification_id = str(data.get('notification_id', '') or '').strip()
        notification = mark_notification_read(notification_id, username)
        if not notification:
            return JsonResponse({'error': 'Notification not found'}, status=404)
        return JsonResponse({
            'ok': True,
            'notification': notification,
            'unread_count': unread_count(username),
        })

    if action == 'read_all':
        changed = mark_all_notifications_read(username)
        return JsonResponse({
            'ok': True,
            'updated': changed,
            'unread_count': unread_count(username),
        })

    return JsonResponse({'error': 'Unknown action'}, status=400)


# ── Hub: monthly schedule ─────────────────────────────────────────────────────
def cmms_hub(request):
    redir = _require_module_access(request, 'activities', 'view')
    if redir:
        return redir
    user = _get_user(request) or {}
    # Pass current month; calendar is rendered in JS
    today = datetime.now()
    return render(request, 'core/cmms_hub.html', _ctx(request, {
        'today': today.strftime('%Y-%m-%d'),
        'current_month': today.strftime('%Y-%m'),
        'can_edit_activities': has_permission(user, 'activities', 'edit'),
        'can_view_permits': has_permission(user, 'permits', 'view'),
    }))


# ── Work page: Excel editor + photos for a specific record ────────────────────
def cmms_handover_legacy(request):
    redir = _require_module_access(request, 'handover', 'view')
    if redir:
        return redir
    return redirect(_default_project_cmms_url(_get_user(request), 'handover/'))


def cmms_work(request, record_id):
    redir = _require_module_access(request, 'activities', 'view')
    if redir:
        return redir
    user = _get_user(request)
    record = get_record(record_id)
    if not record:
        raise Http404('Record not found')
    activity = ensure_activity_checklist(get_activity(record['activity_id']))
    if not activity:
        raise Http404('Activity not found')
    permit = annotate_permit(get_permit_for_record(record_id))
    if not permit and not record.get('completed'):
        permit = annotate_permit(create_or_get_record_permit(record, activity, user or {}))
    if permit and not application_is_active(permit):
        return redirect(f"/cmms/ptw/{permit['id']}/")

    media_url = getattr(__import__('django.conf', fromlist=['settings']).settings, 'MEDIA_URL', '/media/')
    before_urls = [f"{media_url}{p}" for p in record.get('before_photos', [])]
    after_urls  = [f"{media_url}{p}" for p in record.get('after_photos', [])]
    checklist_path = resolve_checklist_path(activity.get('checklist_file', ''))
    activity_type = str(activity.get('type', 'PM') or 'PM').upper()
    sheet_link = get_activity_sheet_link(activity)
    checklist_link = sheet_link if (activity_type == 'CM' or not checklist_path) else ''
    sheet_label = get_activity_sheet_label(activity)

    return render(request, 'core/cmms_work.html', _ctx(request, {
        'record':          record,
        'activity':        activity,
        'before_urls':     before_urls,
        'after_urls':      after_urls,
        'has_checklist':   bool(checklist_path),
        'checklist_link':  checklist_link or '',
        'sheet_label':     sheet_label,
        'media_url':       media_url,
        'permit':          permit,
        'can_edit_activities': has_permission(user or {}, 'activities', 'edit'),
        'can_view_permits': has_permission(user or {}, 'permits', 'view'),
    }))


def cmms_ptw_list(request):
    redir = _require_module_access(request, 'permits', 'view')
    if redir:
        return redir
    user = _get_user(request) or {}
    permits = [annotate_permit(permit) for permit in list_permits() if is_cmms_permit(permit)]
    status_filter = request.GET.get('status', '').strip()
    if status_filter:
        permits = [permit for permit in permits if permit.get('status') == status_filter]
    for permit in permits:
        permit['work_url'] = f"/cmms/work/{permit['record_id']}/" if permit.get('record_id') else ''
        permit['ptw_url'] = f"/cmms/ptw/{permit['id']}/"
        permit['work_accessible'] = bool(permit.get('record_id') and application_is_active(permit))
        permit['can_delete'] = can_delete_permit(permit, user)
    status_counts = {}
    for permit in list_permits():
        if not is_cmms_permit(permit):
            continue
        key = permit.get('status', '')
        status_counts[key] = status_counts.get(key, 0) + 1
    return render(request, 'core/cmms_ptw_list.html', _ctx(request, {
        'permits': permits,
        'status_filter': status_filter,
        'status_counts': status_counts,
        'can_view_activities': has_permission(user, 'activities', 'view'),
        'can_delete_permits': any(permit.get('can_delete') for permit in permits),
    }))


def cmms_ptw_detail(request, permit_id):
    redir = _require_module_access(request, 'permits', 'view')
    if redir:
        return redir
    user = _get_user(request)
    permit = annotate_permit(get_permit(permit_id))
    if not permit:
        raise Http404('Permit not found')
    permit['work_accessible'] = bool(permit.get('record_id') and application_is_active(permit))
    record = get_record(permit.get('record_id', ''))
    activity = ensure_activity_checklist(get_activity(permit.get('activity_id', '')))
    can_permit_edit = has_permission(user or {}, 'permits', 'edit')
    can_application = can_permit_edit and can_edit_application(permit, user)
    can_issuer = can_permit_edit and can_issue_permit(permit, user)
    can_hse = can_permit_edit and can_hse_approve(permit, user)
    can_unlock = can_permit_edit and can_receiver_unlock(permit, user)
    can_close_receiver_flag = can_permit_edit and can_close_receiver(permit, user)
    can_close_issuer_flag = can_permit_edit and can_close_issuer(permit, user)
    can_close_hse_flag = can_permit_edit and can_close_hse(permit, user)
    return render(request, 'core/cmms_ptw.html', _ctx(request, {
        'permit': permit,
        'record': record,
        'activity': activity,
        'can_edit_application': can_application,
        'can_issue_permit': can_issuer,
        'can_hse_approve': can_hse,
        'can_receiver_unlock': can_unlock,
        'can_close_receiver': can_close_receiver_flag,
        'can_close_issuer': can_close_issuer_flag,
        'can_close_hse': can_close_hse_flag,
        'show_close_mode': request.GET.get('mode', '') == 'close',
        'can_view_activities': has_permission(user or {}, 'activities', 'view'),
    }))


def cmms_ptw_download(request, permit_id):
    redir = _require_module_access(request, 'permits', 'view')
    if redir:
        return redir
    permit = annotate_permit(get_permit(permit_id))
    if not permit:
        raise Http404('Permit not found')
    if permit.get('status') == 'closed':
        permit = ensure_final_permit_pdf(permit) or permit
        final_pdf = str((permit or {}).get('final_pdf', '') or '').strip()
        if final_pdf:
            final_pdf_path = Path(getattr(__import__('django.conf', fromlist=['settings']).settings, 'MEDIA_ROOT', '')) / final_pdf
            if final_pdf_path.exists():
                response = HttpResponse(final_pdf_path.read_bytes(), content_type='application/pdf')
                response['Content-Disposition'] = f'attachment; filename="{final_pdf_path.name}"'
                return response
    if permit.get('document_link'):
        return redirect(permit['document_link'])
    buffer = build_permit_docx(permit)
    response = HttpResponse(
        buffer.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )
    response['Content-Disposition'] = f'attachment; filename="{permit_filename(permit)}"'
    return response


def _clean_document_values(payload) -> dict:
    if not isinstance(payload, dict):
        return {}
    cleaned = {}
    for key, value in payload.items():
        clean_key = str(key or '').strip()
        if not clean_key.startswith('t'):
            continue
        cleaned[clean_key] = str(value or '')[:5000]
    return cleaned


def _clean_external_link(value) -> str:
    link = str(value or '').strip()
    if not link:
        return ''
    if link.startswith('http://') or link.startswith('https://'):
        return link
    return ''


@csrf_exempt
def cmms_api_ptw(request, permit_id):
    redir = _require_module_access(request, 'permits', 'view')
    if redir:
        return redir
    user = _get_user(request)
    permit = annotate_permit(get_permit(permit_id))
    if not permit:
        return JsonResponse({'error': 'Permit not found'}, status=404)

    if request.method == 'GET':
        annotated = annotate_permit(permit)
        return JsonResponse({'permit': annotated})

    if request.method != 'POST':
        return JsonResponse({'error': 'Method not allowed'}, status=405)
    if not has_permission(user or {}, 'permits', 'edit'):
        return JsonResponse({'error': 'Permission denied'}, status=403)

    try:
        data = json.loads(request.body)
    except Exception:
        return JsonResponse({'error': 'Invalid JSON'}, status=400)

    action = data.get('action', '').strip()
    now = datetime.now().isoformat()

    if action == 'delete':
        if not can_delete_permit(permit, user):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        deleted = delete_permit(permit_id, (user or {}).get('username', '') or (user or {}).get('name', ''))
        if not deleted:
            return JsonResponse({'error': 'Permit not found'}, status=404)
        return JsonResponse({'ok': True, 'deleted_id': permit_id})

    if action == 'save_document':
        if not can_edit_application(permit, user):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        document_link = _clean_external_link(
            data.get('document_link') or permit.get('document_link') or permit.get('template_link')
        )
        if not document_link:
            return JsonResponse({'error': 'Google Docs permit link is required'}, status=400)
        permit = update_permit(permit_id, {'document_link': document_link})
        return JsonResponse({'ok': True, 'permit': annotate_permit(permit)})

    if action == 'submit_application':
        if not can_edit_application(permit, user):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        document_link = _clean_external_link(data.get('document_link') or permit.get('document_link'))
        if not document_link:
            return JsonResponse({'error': 'Google Docs permit link is required'}, status=400)
        permit = update_permit(permit_id, {
            'document_link': document_link,
            'receiver_name': user.get('name', ''),
            'receiver_username': user.get('username', ''),
            'submitted_at': now,
            'issuer_name': '',
            'issuer_username': '',
            'issuer_signature': '',
            'issued_at': '',
            'hse_name': '',
            'hse_username': '',
            'hse_signature': '',
            'hse_signed_at': '',
            'permit_number': '',
            'receiver_confirmed_permit_number': '',
            'receiver_confirmed_at': '',
            'active_at': '',
            'isolation_cert_number': '',
            'rejection_stage': '',
            'rejection_reason': '',
            'rejected_at': '',
            'rejected_by_name': '',
            'rejected_by_username': '',
            'status': 'pending_issue',
        })
        notify_permit_created(permit, _issuer_notification_users())
        _push_ptw_notifications(
            _usernames_for_users(_issuer_users()),
            title='PTW request needs issuer review',
            message=f"{permit.get('activity_name') or permit.get('equipment') or 'PTW'} was submitted and is waiting for issuer approval.",
            permit=permit,
            focus='issuer-review',
            kind='warning',
            actor_name=user.get('name', ''),
        )
        _push_ptw_notifications(
            [permit.get('receiver_username', '')],
            title='PTW submitted',
            message='Your permit request was submitted and is now waiting for issuer review.',
            permit=permit,
            focus='receiver-application',
            kind='info',
            actor_name=user.get('name', ''),
        )
        return JsonResponse({'ok': True, 'permit': annotate_permit(permit)})

    if action == 'issue':
        if not can_issue_permit(permit, user):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        document_link = _clean_external_link(data.get('document_link') or permit.get('document_link'))
        if not document_link:
            return JsonResponse({'error': 'Google Docs permit link is required'}, status=400)
        payload = {
            'document_link': document_link,
            'issuer_name': user.get('name', ''),
            'issuer_username': user.get('username', ''),
            'issued_at': now,
            'rejection_stage': '',
            'rejection_reason': '',
            'rejected_at': '',
            'rejected_by_name': '',
            'rejected_by_username': '',
            'status': 'pending_hse',
        }
        permit = update_permit(permit_id, payload)
        notify_permit_issued(permit, _hse_notification_users())
        _push_ptw_notifications(
            _usernames_for_users(_hse_users()),
            title='PTW needs HSE approval',
            message=f"{permit.get('activity_name') or permit.get('equipment') or 'PTW'} is waiting for HSE approval.",
            permit=permit,
            focus='hse-review',
            kind='warning',
            actor_name=user.get('name', ''),
        )
        _push_ptw_notifications(
            [permit.get('receiver_username', '')],
            title='Issuer approved the PTW',
            message='Your PTW was reviewed by the issuer and is now waiting for HSE approval.',
            permit=permit,
            focus='permit-status',
            kind='info',
            actor_name=user.get('name', ''),
        )
        return JsonResponse({'ok': True, 'permit': annotate_permit(permit)})

    if action == 'reject_issue':
        if not can_issue_permit(permit, user):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        reason = str(data.get('reason', '') or '').strip()
        if not reason:
            return JsonResponse({'error': 'Rejection reason is required'}, status=400)
        document_link = _clean_external_link(data.get('document_link') or permit.get('document_link'))
        payload = {
            'issuer_name': '',
            'issuer_username': '',
            'issuer_signature': '',
            'issued_at': '',
            'permit_number': '',
            'isolation_cert_number': '',
            'hse_name': '',
            'hse_username': '',
            'hse_signature': '',
            'hse_signed_at': '',
            'rejection_stage': 'issue',
            'rejection_reason': reason,
            'rejected_at': now,
            'rejected_by_name': user.get('name', ''),
            'rejected_by_username': user.get('username', ''),
            'status': 'rejected_by_issuer',
        }
        if document_link:
            payload['document_link'] = document_link
        permit = update_permit(permit_id, payload)
        _push_ptw_notifications(
            _receiver_and_maintenance_usernames(permit.get('receiver_username', '')),
            title='PTW rejected by issuer',
            message=f"Issuer rejected the PTW. Reason: {reason}",
            permit=permit,
            focus='receiver-application',
            kind='error',
            actor_name=user.get('name', ''),
        )
        return JsonResponse({'ok': True, 'permit': annotate_permit(permit)})

    if action == 'hse_approve':
        if not can_hse_approve(permit, user):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        document_link = _clean_external_link(data.get('document_link') or permit.get('document_link'))
        if not document_link:
            return JsonResponse({'error': 'Google Docs permit link is required'}, status=400)
        permit_number = str(data.get('permit_number', '') or '').strip()
        if not permit_number:
            return JsonResponse({'error': 'Permit number is required'}, status=400)
        isolation_number = str(data.get('isolation_cert_number', '') or '').strip()
        permit = update_permit(permit_id, {
            'document_link': document_link,
            'permit_number': permit_number,
            'isolation_cert_number': isolation_number,
            'hse_name': user.get('name', ''),
            'hse_username': user.get('username', ''),
            'hse_signed_at': now,
            'rejection_stage': '',
            'rejection_reason': '',
            'rejected_at': '',
            'rejected_by_name': '',
            'rejected_by_username': '',
            'status': 'pending_receiver_number',
        })
        notify_permit_ready_to_proceed(
            permit,
            _receiver_and_maintenance_emails(permit.get('receiver_username', '')),
        )
        _push_ptw_notifications(
            _receiver_and_maintenance_usernames(permit.get('receiver_username', '')),
            title='PTW approved and ready to proceed',
            message=f"HSE approved the PTW. Enter permit number {permit_number} to proceed.",
            permit=permit,
            focus='receiver-unlock',
            kind='success',
            actor_name=user.get('name', ''),
        )
        return JsonResponse({
            'ok': True,
            'permit': annotate_permit(permit),
            'next_url': f"/cmms/ptw/{permit_id}/",
        })

    if action == 'reject_hse':
        if not can_hse_approve(permit, user):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        reason = str(data.get('reason', '') or '').strip()
        if not reason:
            return JsonResponse({'error': 'Rejection reason is required'}, status=400)
        document_link = _clean_external_link(data.get('document_link') or permit.get('document_link'))
        payload = {
            'permit_number': '',
            'receiver_confirmed_permit_number': '',
            'receiver_confirmed_at': '',
            'isolation_cert_number': '',
            'hse_name': '',
            'hse_username': '',
            'hse_signature': '',
            'hse_signed_at': '',
            'rejection_stage': 'hse',
            'rejection_reason': reason,
            'rejected_at': now,
            'rejected_by_name': user.get('name', ''),
            'rejected_by_username': user.get('username', ''),
            'status': 'rejected_by_hse',
        }
        if document_link:
            payload['document_link'] = document_link
        permit = update_permit(permit_id, payload)
        _push_ptw_notifications(
            _receiver_and_maintenance_usernames(permit.get('receiver_username', '')),
            title='PTW rejected by HSE',
            message=f"HSE rejected the PTW. Reason: {reason}",
            permit=permit,
            focus='receiver-application',
            kind='error',
            actor_name=user.get('name', ''),
        )
        return JsonResponse({'ok': True, 'permit': annotate_permit(permit)})

    if action == 'receiver_unlock':
        if not can_receiver_unlock(permit, user):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        entered_permit_number = str(data.get('entered_permit_number', '') or '').strip()
        issued_permit_number = str(permit.get('permit_number', '') or '').strip()
        if not issued_permit_number:
            return JsonResponse({'error': 'Permit number is not assigned yet'}, status=400)
        if entered_permit_number != issued_permit_number:
            return JsonResponse({'error': 'Permit number does not match the HSE-issued number'}, status=400)
        permit = update_permit(permit_id, {
            'receiver_confirmed_permit_number': entered_permit_number,
            'receiver_confirmed_at': now,
            'active_at': now,
            'status': 'active',
        })
        _push_ptw_notifications(
            [permit.get('receiver_username', '')],
            title='PTW is active',
            message='Permit number confirmed. Work can now proceed on the activity.',
            permit=permit,
            focus='permit-status',
            kind='success',
            actor_name=user.get('name', ''),
        )
        return JsonResponse({
            'ok': True,
            'permit': annotate_permit(permit),
            'next_url': f"/cmms/work/{permit.get('record_id', '')}/",
        })

    if action == 'submit_closure':
        if not can_close_receiver(permit, user):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        document_link = _clean_external_link(data.get('document_link') or permit.get('document_link'))
        if not document_link:
            return JsonResponse({'error': 'Google Docs permit link is required'}, status=400)
        closure_status_text = str(data.get('closure_status_text', '') or '').strip()
        payload = {
            'document_link': document_link,
            'closure_status_text': closure_status_text,
            'closure_requested_at': now,
            'closure_receiver_name': user.get('name', ''),
            'closure_receiver_signed_at': now,
            'closure_issuer_name': '',
            'closure_issuer_signature': '',
            'closure_issuer_signed_at': '',
            'closure_hse_name': '',
            'closure_hse_signature': '',
            'closure_hse_signed_at': '',
            'closure_rejection_stage': '',
            'closure_rejection_reason': '',
            'closure_rejected_at': '',
            'closure_rejected_by_name': '',
            'closure_rejected_by_username': '',
            'status': 'pending_closure_issuer',
        }
        permit = update_permit(permit_id, payload)
        issuer_email = _user_email(permit.get('issuer_username', ''))
        closure_recipients = [issuer_email] if issuer_email else [
            u.get('email', '') for u in _issuer_notification_users()
        ]
        notify_permit_closure_requested(permit, closure_recipients)
        issuer_usernames = [permit.get('issuer_username', '')] if permit.get('issuer_username') else _usernames_for_users(_issuer_users())
        _push_ptw_notifications(
            issuer_usernames,
            title='PTW closure needs issuer review',
            message='The receiver submitted a closure request and it is waiting for issuer review.',
            permit=permit,
            focus='issuer-close-review',
            kind='warning',
            actor_name=user.get('name', ''),
        )
        _push_ptw_notifications(
            [permit.get('receiver_username', '')],
            title='Closure sent to issuer',
            message='Your PTW closure request was submitted and is now waiting for issuer review.',
            permit=permit,
            focus='receiver-close',
            kind='info',
            actor_name=user.get('name', ''),
        )
        return JsonResponse({'ok': True, 'permit': annotate_permit(permit)})

    if action == 'issuer_close':
        if not can_close_issuer(permit, user):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        document_link = _clean_external_link(data.get('document_link') or permit.get('document_link'))
        if not document_link:
            return JsonResponse({'error': 'Google Docs permit link is required'}, status=400)
        payload = {
            'document_link': document_link,
            'closure_issuer_name': user.get('name', ''),
            'closure_issuer_signed_at': now,
            'closure_rejection_stage': '',
            'closure_rejection_reason': '',
            'closure_rejected_at': '',
            'closure_rejected_by_name': '',
            'closure_rejected_by_username': '',
            'status': 'pending_closure_hse',
        }
        permit = update_permit(permit_id, payload)
        notify_permit_closure_hse_required(permit, _hse_notification_users())
        _push_ptw_notifications(
            _usernames_for_users(_hse_users()),
            title='PTW closure needs HSE approval',
            message='The issuer reviewed the closure request and it is now waiting for HSE closure.',
            permit=permit,
            focus='hse-close-review',
            kind='warning',
            actor_name=user.get('name', ''),
        )
        _push_ptw_notifications(
            [permit.get('receiver_username', '')],
            title='Closure sent to HSE',
            message='The issuer reviewed your closure request. It is now waiting for HSE closure.',
            permit=permit,
            focus='permit-status',
            kind='info',
            actor_name=user.get('name', ''),
        )
        return JsonResponse({'ok': True, 'permit': annotate_permit(permit)})

    if action == 'reject_closure_issuer':
        if not can_close_issuer(permit, user):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        reason = str(data.get('reason', '') or '').strip()
        if not reason:
            return JsonResponse({'error': 'Rejection reason is required'}, status=400)
        document_link = _clean_external_link(data.get('document_link') or permit.get('document_link'))
        payload = {
            'closure_issuer_name': '',
            'closure_issuer_signature': '',
            'closure_issuer_signed_at': '',
            'closure_hse_name': '',
            'closure_hse_signature': '',
            'closure_hse_signed_at': '',
            'closure_rejection_stage': 'issuer',
            'closure_rejection_reason': reason,
            'closure_rejected_at': now,
            'closure_rejected_by_name': user.get('name', ''),
            'closure_rejected_by_username': user.get('username', ''),
            'status': 'closure_rejected_by_issuer',
        }
        if document_link:
            payload['document_link'] = document_link
        permit = update_permit(permit_id, payload)
        _push_ptw_notifications(
            _receiver_and_maintenance_usernames(permit.get('receiver_username', '')),
            title='PTW closure rejected by issuer',
            message=f"Issuer rejected the closure request. Reason: {reason}",
            permit=permit,
            focus='receiver-close',
            kind='error',
            actor_name=user.get('name', ''),
        )
        return JsonResponse({'ok': True, 'permit': annotate_permit(permit)})

    if action == 'hse_close':
        if not can_close_hse(permit, user):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        document_link = _clean_external_link(data.get('document_link') or permit.get('document_link'))
        if not document_link:
            return JsonResponse({'error': 'Google Docs permit link is required'}, status=400)
        payload = {
            'document_link': document_link,
            'closure_hse_name': user.get('name', ''),
            'closure_hse_signed_at': now,
            'closure_rejection_stage': '',
            'closure_rejection_reason': '',
            'closure_rejected_at': '',
            'closure_rejected_by_name': '',
            'closure_rejected_by_username': '',
            'closed_at': now,
            'status': 'closed',
        }
        permit = update_permit(permit_id, payload)
        permit = ensure_final_permit_pdf(permit) or permit
        if permit.get('record_id'):
            update_record(permit['record_id'], {
                'completed': True,
                'completed_at': now,
            })
        close_emails = _closure_notification_emails(permit)
        notify_permit_closed(permit, close_emails)
        _push_ptw_notifications(
            _closure_notification_usernames(permit),
            title='PTW closed',
            message='The permit was fully closed by HSE.',
            permit=permit,
            focus='permit-summary',
            kind='success',
            actor_name=user.get('name', ''),
        )
        return JsonResponse({
            'ok': True,
            'permit': annotate_permit(permit),
            'next_url': f"/cmms/work/{permit.get('record_id', '')}/",
        })

    if action == 'reject_closure_hse':
        if not can_close_hse(permit, user):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        reason = str(data.get('reason', '') or '').strip()
        if not reason:
            return JsonResponse({'error': 'Rejection reason is required'}, status=400)
        document_link = _clean_external_link(data.get('document_link') or permit.get('document_link'))
        payload = {
            'closure_hse_name': '',
            'closure_hse_signature': '',
            'closure_hse_signed_at': '',
            'closure_rejection_stage': 'hse',
            'closure_rejection_reason': reason,
            'closure_rejected_at': now,
            'closure_rejected_by_name': user.get('name', ''),
            'closure_rejected_by_username': user.get('username', ''),
            'status': 'closure_rejected_by_hse',
        }
        if document_link:
            payload['document_link'] = document_link
        permit = update_permit(permit_id, payload)
        _push_ptw_notifications(
            _receiver_and_maintenance_usernames(permit.get('receiver_username', '')),
            title='PTW closure rejected by HSE',
            message=f"HSE rejected the closure request. Reason: {reason}",
            permit=permit,
            focus='receiver-close',
            kind='error',
            actor_name=user.get('name', ''),
        )
        return JsonResponse({'ok': True, 'permit': annotate_permit(permit)})

    return JsonResponse({'error': 'Unknown action'}, status=400)


# ── API: Activities schedule ──────────────────────────────────────────────────
@csrf_exempt
def cmms_api_activities(request):
    """
    GET  ?month=YYYY-MM  → list activities for that month
    POST                 → create a new scheduled activity
    """
    redir = _require_module_access(request, 'activities', 'view')
    if redir:
        return redir
    user = _get_user(request)

    if request.method == 'GET':
        month = request.GET.get('month', datetime.now().strftime('%Y-%m'))
        activities = []
        for activity in get_activities_for_month(month):
            activity = ensure_activity_checklist(activity)
            record = get_record_for_activity_date(activity.get('id', ''), activity.get('scheduled_date', ''))
            permit = annotate_permit(get_permit_for_record(record.get('id', ''))) if record else None
            status = 'planned'
            if record:
                status = 'completed' if record.get('completed') else 'in_progress'
            permit_required = bool(
                record and not record.get('completed') and (not permit or not application_is_active(permit))
            )
            activities.append({
                **activity,
                'record_id': record.get('id') if record else '',
                'record_started': bool(record),
                'record_completed': bool(record and record.get('completed')),
                'record_completed_at': record.get('completed_at') if record else None,
                'record_status': status,
                'record_url': f"/cmms/work/{record.get('id')}/" if record else '',
                'permit_id': permit.get('id') if permit else '',
                'permit_status': permit.get('status') if permit else '',
                'permit_status_label': permit.get('status_label') if permit else '',
                'permit_number': permit.get('permit_number') if permit else '',
                'permit_required': permit_required,
                'permit_url': f"/cmms/ptw/{permit.get('id')}/" if permit else '',
            })
        return JsonResponse({'activities': activities})

    if request.method == 'POST':
        if not has_permission(user or {}, 'activities', 'edit'):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        try:
            data = json.loads(request.body)
        except Exception:
            return JsonResponse({'error': 'Invalid JSON'}, status=400)
        # Derive month from scheduled_date
        scheduled_date = data.get('scheduled_date', '')
        if scheduled_date and len(scheduled_date) >= 7:
            data['month'] = scheduled_date[:7]
        data['created_by'] = user.get('username', '')
        activity = create_activity(data)
        return JsonResponse({'activity': activity}, status=201)

    if request.method == 'DELETE':
        if not has_permission(user or {}, 'activities', 'edit'):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        try:
            data = json.loads(request.body)
        except Exception:
            return JsonResponse({'error': 'Invalid JSON'}, status=400)
        activity_id = data.get('id', '')
        if delete_activity(activity_id):
            return JsonResponse({'ok': True})
        return JsonResponse({'error': 'Not found'}, status=404)

    return JsonResponse({'error': 'Method not allowed'}, status=405)


@csrf_exempt
def cmms_api_activity(request, activity_id):
    """PATCH/DELETE for a single activity."""
    redir = _require_module_access(request, 'activities', 'edit')
    if redir:
        return redir
    user = _get_user(request)

    if request.method == 'PATCH':
        try:
            data = json.loads(request.body)
        except Exception:
            return JsonResponse({'error': 'Invalid JSON'}, status=400)
        activity = update_activity(activity_id, data)
        if not activity:
            return JsonResponse({'error': 'Not found'}, status=404)
        return JsonResponse({'activity': activity})

    if request.method == 'DELETE':
        try:
            data = json.loads(request.body) if request.body else {}
        except Exception:
            return JsonResponse({'error': 'Invalid JSON'}, status=400)
        scheduled_date = str(data.get('scheduled_date', '') or '').strip()
        deleted = (
            delete_activity_occurrence(activity_id, scheduled_date)
            if scheduled_date else
            delete_activity(activity_id)
        )
        if deleted:
            return JsonResponse({'ok': True})
        return JsonResponse({'error': 'Not found'}, status=404)

    return JsonResponse({'error': 'Method not allowed'}, status=405)


# ── API: Start / get record ───────────────────────────────────────────────────
@csrf_exempt
def cmms_api_start(request):
    """
    POST {activity_id, date} → start (or resume) a record, return record_id
    """
    redir = _require_module_access(request, 'activities', 'edit')
    if redir:
        return redir
    user = _get_user(request)

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
        except Exception:
            return JsonResponse({'error': 'Invalid JSON'}, status=400)
        activity_id = data.get('activity_id', '')
        date = data.get('date', datetime.now().strftime('%Y-%m-%d'))
        selected_permit_name = str(data.get('selected_permit_name', '') or '').strip()
        selected_permit_link = str(data.get('selected_permit_link', '') or '').strip()
        if not activity_id:
            return JsonResponse({'error': 'activity_id required'}, status=400)
        activity = ensure_activity_checklist(get_activity(activity_id))
        if not activity:
            return JsonResponse({'error': 'Activity not found'}, status=404)
        if not activity_occurs_on_date(activity, date):
            return JsonResponse({'error': 'Activity can only be started on its planned schedule'}, status=400)
        permit_options = get_activity_permit_options(activity.get('name', ''))
        if permit_options and not selected_permit_name:
            return JsonResponse({'error': 'Permit selection required'}, status=400)
        if selected_permit_name and not selected_permit_link:
            normalized_name = selected_permit_name.strip().lower()
            for option in permit_options:
                if str(option.get('name', '')).strip().lower() == normalized_name:
                    selected_permit_link = str(option.get('link', '') or '').strip()
                    break
        record = start_record(activity_id, date, user.get('username', ''), user.get('name', ''))
        permit = create_or_get_record_permit(
            record,
            activity,
            user,
            permit_name=selected_permit_name,
            permit_link=selected_permit_link,
        )
        next_url = f"/cmms/work/{record['id']}/" if application_is_active(permit) else f"/cmms/ptw/{permit['id']}/"
        return JsonResponse({
            'record_id': record['id'],
            'permit_id': permit.get('id'),
            'permit_status': permit.get('status'),
            'next_url': next_url,
            'permit_required': not application_is_active(permit),
        })

    return JsonResponse({'error': 'POST only'}, status=405)


# ── API: Excel checklist (GET parse + saved values, POST save) ────────────────
@csrf_exempt
def cmms_api_excel(request, record_id):
    redir = _require_module_access(request, 'activities', 'view')
    if redir:
        return redir

    record = get_record(record_id)
    if not record:
        return JsonResponse({'error': 'Record not found'}, status=404)
    activity = ensure_activity_checklist(get_activity(record['activity_id']))
    if not activity:
        return JsonResponse({'error': 'Activity not found'}, status=404)

    if request.method == 'GET':
        checklist_path = resolve_checklist_path(activity.get('checklist_file', ''))
        if not checklist_path:
            return JsonResponse({'sheets': [], 'values': {}})
        if not checklist_path.exists():
            return JsonResponse({'error': 'Checklist file not found'}, status=404)
        try:
            sheets = parse_excel_checklist(checklist_path)
        except Exception as e:
            return JsonResponse({'error': f'Parse error: {e}'}, status=500)
        return JsonResponse({'sheets': sheets, 'values': record.get('excel_values', {})})

    if request.method == 'POST':
        user = _get_user(request)
        if not has_permission(user or {}, 'activities', 'edit'):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        try:
            data = json.loads(request.body)
        except Exception:
            return JsonResponse({'error': 'Invalid JSON'}, status=400)
        update_record(record_id, {'excel_values': data.get('excel_values', {})})
        return JsonResponse({'ok': True})

    return JsonResponse({'error': 'GET or POST only'}, status=405)


# ── API: Photo upload / delete ────────────────────────────────────────────────
@csrf_exempt
def cmms_api_photos(request, record_id):
    redir = _require_module_access(request, 'activities', 'view')
    if redir:
        return redir

    record = get_record(record_id)
    if not record:
        return JsonResponse({'error': 'Record not found'}, status=404)

    if request.method == 'POST':
        user = _get_user(request)
        if not has_permission(user or {}, 'activities', 'edit'):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        phase = request.POST.get('phase', 'before')
        if phase not in ('before', 'after'):
            return JsonResponse({'error': 'phase must be before or after'}, status=400)
        uploaded = request.FILES.getlist('photos')
        if not uploaded:
            return JsonResponse({'error': 'No files uploaded'}, status=400)
        media_url = getattr(__import__('django.conf', fromlist=['settings']).settings, 'MEDIA_URL', '/media/')
        saved = []
        for f in uploaded:
            rel = save_photo(record_id, phase, f)
            saved.append({'rel': rel, 'url': f"{media_url}{rel}"})
        return JsonResponse({'saved': saved})

    if request.method == 'DELETE':
        user = _get_user(request)
        if not has_permission(user or {}, 'activities', 'edit'):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        try:
            data = json.loads(request.body)
        except Exception:
            return JsonResponse({'error': 'Invalid JSON'}, status=400)
        phase = data.get('phase', 'before')
        rel   = data.get('rel', '')
        if delete_photo(record_id, phase, rel):
            return JsonResponse({'ok': True})
        return JsonResponse({'error': 'Not found'}, status=404)

    return JsonResponse({'error': 'POST or DELETE only'}, status=405)


# ── API: Mark complete ────────────────────────────────────────────────────────
@csrf_exempt
def cmms_api_complete(request, record_id):
    redir = _require_module_access(request, 'activities', 'edit')
    if redir:
        return redir
    if request.method == 'POST':
        record = get_record(record_id)
        if not record:
            return JsonResponse({'error': 'Not found'}, status=404)
        permit = annotate_permit(get_permit_for_record(record_id))
        if not permit:
            return JsonResponse({'error': 'Permit must be completed before closing the activity'}, status=400)
        if permit.get('status') != 'closed':
            return JsonResponse({'error': 'Permit is not closed yet. Complete PTW closure first.'}, status=400)
        update_record(record_id, {
            'completed':    True,
            'completed_at': datetime.now().isoformat(),
        })
        return JsonResponse({'ok': True})
    return JsonResponse({'error': 'POST only'}, status=405)


# ── Download ZIP ──────────────────────────────────────────────────────────────
def cmms_download_zip(request, record_id):
    redir = _require_module_access(request, 'activities', 'view')
    if redir:
        return redir
    buf = generate_zip(record_id)
    if not buf:
        raise Http404('Record not found or no data')
    record = get_record(record_id)
    fname = f"{record.get('activity_name','report')}_{record.get('date','')}.zip".replace(' ', '_')
    resp = HttpResponse(buf, content_type='application/zip')
    resp['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


# ── API: Available checklist files ────────────────────────────────────────────
def cmms_api_checklists(request):
    redir = _require_module_access(request, 'activities', 'view')
    if redir:
        return redir
    user = _get_user(request)

    if request.method == 'GET':
        return JsonResponse({'files': get_checklist_files()})

    if request.method == 'POST':
        if not has_permission(user or {}, 'activities', 'edit'):
            return JsonResponse({'error': 'Permission denied'}, status=403)
        uploaded = request.FILES.get('checklist')
        if not uploaded:
            return JsonResponse({'error': 'Checklist file is required'}, status=400)
        try:
            file_info = save_checklist_file(uploaded)
        except ValueError as exc:
            return JsonResponse({'error': str(exc)}, status=400)
        return JsonResponse({'file': file_info}, status=201)


# ── API: Activity list from live Google Sheet ─────────────────────────────────
def cmms_api_checklist_activities(request):
    """
    GET /api/cmms/checklist-activities/
    Returns all activity mappings from the configured live Google Sheet,
    including PM checklist links, PTW options, and CM report links.
    """
    redir = _require_module_access(request, 'activities', 'view')
    if redir:
        return redir
    return JsonResponse({'activities': get_all_checklist_activities()})

    return JsonResponse({'error': 'Method not allowed'}, status=405)
