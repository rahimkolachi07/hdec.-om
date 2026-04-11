"""
CMMS Views — Activities, Checklists, Permits, PDF, ZIP
"""
import json, os
from datetime import datetime
from pathlib import Path
from django.shortcuts import render, redirect
from django.http import JsonResponse, HttpResponse, Http404
from django.views.decorators.csrf import csrf_exempt
from django.conf import settings

from .auth_utils import (
    get_all_users, get_users_by_role, get_user_detail, ROLE_LABELS,
    normalize_user_state, has_permission,
    can_access_country, can_access_project,
)
from .project_utils import get_countries
from .cmms_utils import (
    WORK_TYPES,
    get_activities, get_activity, get_activities_for_user,
    create_activity, update_activity, delete_activity,
    # Records
    get_records, get_record, get_or_create_record,
    get_records_for_activity, update_record, save_photo,
    generate_record_zip, generate_activity_month_zip,
    # Permits
    get_permits, get_permit, get_permits_for_user,
    create_permit, update_permit,
    get_next_permit_number, get_next_isolation_number,
    generate_permit_pdf, generate_icc_pdf, generate_permit_docx,
    PERMIT_STATUSES, CHECKLISTS_DIR, MEDIA_ROOT,
    # Manpower
    get_duty_staff,
    # Daily auto-records
    auto_create_daily_records, get_today_records_for_user,
    # Smart dashboard
    get_today_tasks_for_dashboard,
    # Schedule import
    import_activities_from_schedule,
    # Technicians from schedule
    get_all_technicians_from_schedule,
    # Handover / Shift Log
    get_handovers, get_handover, get_handovers_by_date, get_handover_dates,
    create_handover, update_handover, save_handover_image,
)
from .email_utils import (
    notify_permit_created, notify_permit_issued,
    notify_permit_approved, notify_activity_assigned,
)


def _get_user(request):
    user = normalize_user_state(request.session.get('hdec_user'))
    if user != request.session.get('hdec_user'):
        request.session['hdec_user'] = user
        request.session.modified = True
    return user


def _ctx(request, extra=None):
    user = _get_user(request)
    ctx = {
        'current_user': user or {},
        'is_admin': bool(user and user.get('role') == 'admin'),
    }
    if extra:
        ctx.update(extra)
    return ctx


def _require_login(request):
    if not _get_user(request):
        return redirect('/login/')
    return None


def _has_any_project_access(user):
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


def _require_module(request, module_id=None, level='view', api=False, any_of=None):
    user = _get_user(request)
    if not user:
        if api:
            return JsonResponse({'error': 'Login required'}, status=401)
        return redirect('/login/')
    if not _has_any_project_access(user):
        if api:
            return JsonResponse({'error': 'Forbidden'}, status=403)
        return redirect('/')
    allowed = True
    if any_of:
        allowed = any(has_permission(user, mid, level) for mid in any_of)
    elif module_id:
        allowed = has_permission(user, module_id, level)
    if not allowed:
        if api:
            return JsonResponse({'error': 'Forbidden'}, status=403)
        return redirect('/')
    return None


def _require_roles(request, *roles):
    user = _get_user(request)
    if not user:
        return redirect('/login/')
    if user.get('role') not in roles:
        return HttpResponse('Forbidden — insufficient role', status=403)
    return None


# ── CMMS HUB ─────────────────────────────────────────────────────────────

def cmms_hub(request):
    redir = _require_login(request)
    if redir:
        return redir
    redir = _require_module(request, any_of=('activities', 'permits', 'handover'))
    if redir:
        return redir
    user = _get_user(request)
    role = user.get('role', '')
    today = datetime.now().strftime('%Y-%m-%d')

    # Auto-create records for duty staff
    auto_create_daily_records(today)

    activities  = get_activities_for_user(user['username'], role)
    records     = get_records()
    permits     = get_permits_for_user(user['username'], role)

    # Smart today's task list (scheduling engine)
    today_tasks = get_today_tasks_for_dashboard(today)

    completed_records  = [r for r in records if r.get('completed')]
    active_permits     = [p for p in permits if p.get('status') in ('active', 'waiting_for_close')]
    pending_permits    = [p for p in permits if p.get('status') in ('pending_issue', 'pending_hse')]

    # Counts for stat cards
    tasks_done    = sum(1 for t in today_tasks if t['status'] == 'done')
    tasks_pending = sum(1 for t in today_tasks if t['status'] == 'pending')
    tasks_total   = len(today_tasks)

    from datetime import date as _date
    _td = _date.today()
    today_weekday  = _td.strftime('%A')
    today_date_str = _td.strftime('%d %B %Y')

    # Handover stats
    today_handovers = get_handovers_by_date(today)
    today_handover_count = sum(1 for v in today_handovers.values() if v)

    # today_records used in template (tasks list for current user)
    today_records = get_today_records_for_user(user['username'], role, today)

    ctx = _ctx(request, {
        'today':           today,
        'today_weekday':   today_weekday,
        'today_date_str':  today_date_str,
        'today_tasks':     today_tasks,
        'tasks_total':     tasks_total,
        'tasks_done':      tasks_done,
        'tasks_pending':   tasks_pending,
        'total_activities': len(activities),
        'total_records':   len(records),
        'completed_records': len(completed_records),
        'total_permits':   len(permits),
        'active_permits':  len(active_permits),
        'pending_permits': len(pending_permits),
        'recent_permits':  sorted(permits, key=lambda p: p.get('created_at', ''), reverse=True)[:5],
        'today_records':   today_records,
        'today_handovers': today_handovers,
        'today_handover_count': today_handover_count,
    })
    return render(request, 'core/cmms_hub.html', ctx)


# ── ACTIVITIES ────────────────────────────────────────────────────────────

def cmms_activities(request):
    redir = _require_login(request)
    if redir:
        return redir
    redir = _require_module(request, 'activities', 'view')
    if redir:
        return redir
    user = _get_user(request)
    role = user.get('role', '')
    today = datetime.now().strftime('%Y-%m-%d')

    month_filter = request.GET.get('month', datetime.now().strftime('%Y-%m'))

    # Auto-create today's records whenever this page is visited
    auto_create_daily_records(today)

    all_acts = get_activities_for_user(user['username'], role)
    activities = [a for a in all_acts if a.get('month', '').startswith(month_filter)] if month_filter else all_acts

    # Enrich with record completion status and today's record
    for a in activities:
        recs = get_records_for_activity(a['id'])
        a['record_count'] = len(recs)
        a['completed_count'] = len([r for r in recs if r.get('completed')])
        # Find today's record for this user
        today_rec = next((r for r in recs
                          if r['date'] == today and
                          (role == 'admin' or r.get('engineer') == user['username']
                           or r.get('technician') == user['username'])), None)
        a['has_today_record'] = today_rec is not None
        a['today_record_id'] = today_rec['id'] if today_rec else None
        a['today_completed'] = today_rec.get('completed', False) if today_rec else False

    engineers   = get_users_by_role('maintenance_engineer')
    # Technicians come from schedule_store (all 19), on-duty ones listed first
    technicians = get_all_technicians_from_schedule(today)

    ctx = _ctx(request, {
        'activities':   activities,
        'month_filter': month_filter,
        'today':        today,
        'engineers':    engineers,
        'technicians':  technicians,
        'role_labels':  ROLE_LABELS,
    })
    return render(request, 'core/cmms_activities.html', ctx)


@csrf_exempt
def cmms_activity_api(request):
    """Admin: create / delete activities."""
    user = _get_user(request)
    if not user or user.get('role') != 'admin':
        return JsonResponse({'error': 'Admin only'}, status=403)

    if request.method == 'POST':
        # Supports multipart (with file upload) or JSON
        if request.content_type and 'multipart' in request.content_type:
            data = request.POST
            checklist_file_path = ''
            if 'checklist_file' in request.FILES:
                f = request.FILES['checklist_file']
                ext = Path(f.name).suffix.lower()
                safe_name = f"checklist_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}"
                dest = CHECKLISTS_DIR / safe_name
                with open(dest, 'wb') as out:
                    for chunk in f.chunks():
                        out.write(chunk)
                checklist_file_path = f"cmms/checklists/{safe_name}"

            # Parse checklist items from JSON string field
            raw_items = data.get('checklist_items', '[]')
            try:
                checklist_items = json.loads(raw_items)
            except Exception:
                checklist_items = []

            scheduled_date = data.get('scheduled_date', '')
            # Derive month from scheduled_date if provided, otherwise use month field
            month = scheduled_date[:7] if scheduled_date else data.get('month', '')
            act = create_activity({
                'month':               month,
                'scheduled_date':      scheduled_date,
                'name':                data.get('name', ''),
                'equipment':           data.get('equipment', ''),
                'location':            data.get('location', ''),
                'frequency':           data.get('frequency', 'once'),
                'assigned_engineer':   data.get('assigned_engineer', ''),
                'assigned_technician': data.get('assigned_technician', ''),
                'checklist_items':     checklist_items,
                'checklist_file':      checklist_file_path,
                'notes':               data.get('notes', ''),
                'created_by':          user['username'],
            })
            # Send email to assigned engineer
            eng = get_user_detail(act.get('assigned_engineer', ''))
            if eng and eng.get('email'):
                notify_activity_assigned(act, eng['email'])

            return JsonResponse({'ok': True, 'activity': act})
        else:
            try:
                data = json.loads(request.body)
            except Exception:
                return JsonResponse({'error': 'Invalid JSON'}, status=400)

            action = data.get('action')
            if action == 'delete':
                delete_activity(data.get('id', ''))
                return JsonResponse({'ok': True})
            if action == 'import_schedule':
                month = data.get('month', datetime.now().strftime('%Y-%m'))
                try:
                    result = import_activities_from_schedule(month)
                    return JsonResponse({'ok': True,
                                        'created': len(result['created']),
                                        'skipped': result['skipped']})
                except FileNotFoundError as e:
                    return JsonResponse({'error': str(e)}, status=404)
                except Exception as e:
                    return JsonResponse({'error': str(e)}, status=500)
            if action == 'reset_all':
                # Delete all activities and all records
                from .cmms_utils import _save, ACTIVITIES_FILE, RECORDS_FILE
                _save(ACTIVITIES_FILE, [])
                _save(RECORDS_FILE, [])
                return JsonResponse({'ok': True, 'message': 'All activities and records cleared'})
            if action == 'update':
                act = update_activity(data.get('id', ''), data.get('updates', {}))
                return JsonResponse({'ok': bool(act), 'activity': act})

    return JsonResponse({'error': 'Bad request'}, status=400)


def cmms_activity_detail(request, activity_id):
    redir = _require_login(request)
    if redir:
        return redir
    redir = _require_module(request, 'activities', 'view')
    if redir:
        return redir
    user = _get_user(request)
    role = user.get('role', '')

    activity = get_activity(activity_id)
    if not activity:
        raise Http404('Activity not found')

    # Determine viewing date
    date = request.GET.get('date', datetime.now().strftime('%Y-%m-%d'))

    # Get or create a record for this engineer+date (only for engineers)
    record = None
    if role in ('maintenance_engineer', 'admin'):
        record = get_or_create_record(
            activity_id, date,
            user['username'], user['name']
        )
    elif role == 'technician':
        # Technician: find record for today
        recs = get_records_for_activity(activity_id)
        record = next((r for r in recs if r['date'] == date), None)

    all_records = get_records_for_activity(activity_id)

    # Attach photo URL prefixes
    media_url = getattr(settings, 'MEDIA_URL', '/media/')
    if record:
        record['before_photo_urls'] = [
            f"{media_url}{p}" for p in record.get('before_photos', [])
        ]
        record['after_photo_urls'] = [
            f"{media_url}{p}" for p in record.get('after_photos', [])
        ]

    ctx = _ctx(request, {
        'activity': activity,
        'record': record,
        'all_records': all_records,
        'selected_date': date,
        'media_url': media_url,
    })
    return render(request, 'core/cmms_activity_detail.html', ctx)


@csrf_exempt
def cmms_checklist_api(request, record_id):
    """Engineer: submit checklist items + signature."""
    redir = _require_module(request, 'activities', 'view', api=True)
    if redir:
        return redir
    user = _get_user(request)
    if not user:
        return JsonResponse({'error': 'Login required'}, status=401)
    if user.get('role') not in ('maintenance_engineer', 'admin'):
        return JsonResponse({'error': 'Engineer role required'}, status=403)

    record = get_record(record_id)
    if not record:
        return JsonResponse({'error': 'Record not found'}, status=404)

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
        except Exception:
            return JsonResponse({'error': 'Invalid JSON'}, status=400)

        updates = {}
        if 'checkpoints' in data:
            updates['checkpoints'] = data['checkpoints']
        if 'remarks' in data:
            updates['remarks'] = data['remarks']
        if 'engineer_signature' in data:
            updates['engineer_signature'] = data['engineer_signature']
        if data.get('complete'):
            updates['completed'] = True
            updates['completed_at'] = datetime.now().isoformat()

        updated = update_record(record_id, updates)
        return JsonResponse({'ok': True, 'record': updated})

    return JsonResponse({'error': 'POST only'}, status=405)


@csrf_exempt
def cmms_photo_api(request, record_id):
    """Technician / engineer: upload before/after photos."""
    redir = _require_module(request, 'activities', 'view', api=True)
    if redir:
        return redir
    user = _get_user(request)
    if not user:
        return JsonResponse({'error': 'Login required'}, status=401)

    record = get_record(record_id)
    if not record:
        return JsonResponse({'error': 'Record not found'}, status=404)

    if request.method == 'POST':
        phase = request.POST.get('phase', 'before')  # 'before' or 'after'
        if phase not in ('before', 'after'):
            return JsonResponse({'error': 'phase must be before or after'}, status=400)

        uploaded = []
        for f in request.FILES.getlist('photos'):
            # Validate file type
            ext = Path(f.name).suffix.lower()
            if ext not in ('.jpg', '.jpeg', '.png', '.webp', '.heic'):
                continue
            rel_path = save_photo(record_id, phase, f)
            uploaded.append(rel_path)

        if uploaded:
            key = f'{phase}_photos'
            existing = record.get(key, [])
            update_record(record_id, {key: existing + uploaded})

        media_url = getattr(settings, 'MEDIA_URL', '/media/')
        return JsonResponse({
            'ok': True,
            'uploaded': uploaded,
            'urls': [f"{media_url}{p}" for p in uploaded],
        })

    if request.method == 'DELETE':
        data = json.loads(request.body)
        phase = data.get('phase', 'before')
        filename = data.get('filename', '')
        # Remove from record
        key = f'{phase}_photos'
        existing = record.get(key, [])
        updated_list = [p for p in existing if not p.endswith(filename)]
        update_record(record_id, {key: updated_list})
        return JsonResponse({'ok': True})

    return JsonResponse({'error': 'Bad request'}, status=400)


def cmms_download_zip(request, record_id):
    """Download ZIP for a single record (one day's activity)."""
    redir = _require_login(request)
    if redir:
        return redir
    redir = _require_module(request, 'activities', 'view')
    if redir:
        return redir

    buf = generate_record_zip(record_id)
    if not buf:
        raise Http404('Record not found or no data')

    record = get_record(record_id)
    filename = f"activity_{record.get('activity_name','record')}_{record.get('date','')}.zip".replace(' ', '_')
    resp = HttpResponse(buf.read(), content_type='application/zip')
    resp['Content-Disposition'] = f'attachment; filename="{filename}"'
    return resp


def cmms_download_activity_zip(request, activity_id):
    """Download combined ZIP for all records of an activity."""
    redir = _require_login(request)
    if redir:
        return redir
    redir = _require_module(request, 'activities', 'view')
    if redir:
        return redir

    buf = generate_activity_month_zip(activity_id)
    if not buf:
        raise Http404('Activity not found or no data')

    activity = get_activity(activity_id)
    filename = f"activity_{activity['name']}_{activity['month']}.zip".replace(' ', '_')
    resp = HttpResponse(buf.read(), content_type='application/zip')
    resp['Content-Disposition'] = f'attachment; filename="{filename}"'
    return resp


# ── PERMITS ───────────────────────────────────────────────────────────────

def cmms_permits(request):
    redir = _require_login(request)
    if redir:
        return redir
    redir = _require_module(request, 'permits', 'view')
    if redir:
        return redir
    user = _get_user(request)
    role = user.get('role', '')

    status_filter = request.GET.get('status', '')
    permits = get_permits_for_user(user['username'], role)

    if status_filter:
        permits = [p for p in permits if p.get('status') == status_filter]

    permits = sorted(permits, key=lambda p: p.get('created_at', ''), reverse=True)

    # Enrich with labels
    work_types = {
        'cold_work': 'Cold Work', 'hot_work': 'Hot Work',
        'electrical': 'Electrical', 'confined_space': 'Confined Space',
        'mechanical': 'Mechanical', 'civil': 'Civil', 'at_height': 'At Height',
    }
    for p in permits:
        p['status_label'] = PERMIT_STATUSES.get(p.get('status', ''), p.get('status', ''))
        p['work_type_label'] = work_types.get(p.get('work_type', ''), p.get('work_type', ''))

    ctx = _ctx(request, {
        'permits': permits,
        'status_filter': status_filter,
        'permit_statuses': PERMIT_STATUSES,
        'can_create': role in ('maintenance_engineer', 'admin'),
    })
    return render(request, 'core/cmms_permits.html', ctx)


def cmms_permit_detail(request, permit_id=None):
    redir = _require_login(request)
    if redir:
        return redir
    redir = _require_module(request, 'permits', 'view')
    if redir:
        return redir
    user = _get_user(request)
    role = user.get('role', '')

    permit = get_permit(permit_id) if permit_id else None

    work_types = {
        'cold_work': 'Cold Work', 'hot_work': 'Hot Work',
        'electrical': 'Electrical Work', 'confined_space': 'Confined Space Entry',
        'mechanical': 'Mechanical Work', 'civil': 'Civil Work', 'at_height': 'Work at Height',
    }

    precaution_items = [
        {'key': 'safe_distance',     'label': 'Safe Working Distance',          'default': 'N/A'},
        {'key': 'loto_required',     'label': 'LOTO / Isolation Required',       'default': 'No'},
        {'key': 'confined_space',    'label': 'Confined Space Entry',            'default': 'No'},
        {'key': 'power_isolated',    'label': 'Power Isolated',                  'default': 'Yes'},
        {'key': 'lines_de_energized','label': 'Lines / Equipment De-Energized',  'default': 'Yes'},
        {'key': 'tools_tested',      'label': 'Tools / Instruments Tested',      'default': 'Yes'},
    ]

    sig_triples = []
    if permit:
        sig_triples = [
            ('Receiver',    permit.get('receiver_name'), permit.get('receiver_signature'), permit.get('created_at')),
            ('Issuer',      permit.get('issuer_name'),   permit.get('issuer_signature'),   permit.get('issued_at')),
            ('HSE Officer', permit.get('hse_name'),      permit.get('hse_signature'),      permit.get('hse_signed_at')),
        ]

    ctx = _ctx(request, {
        'permit': permit,
        'work_types': work_types,
        'permit_statuses': PERMIT_STATUSES,
        'role': role,
        'can_issue':     role in ('operation_engineer', 'admin'),
        'can_hse_sign':  role in ('hse_engineer', 'admin'),
        'can_close':     role in ('operation_engineer', 'admin', 'hse_engineer'),
        'precaution_items': precaution_items,
        'sig_triples':      sig_triples,
    })
    return render(request, 'core/cmms_permit_detail.html', ctx)


@csrf_exempt
def cmms_permit_api(request):
    """Create permit / workflow actions via JSON API."""
    redir = _require_module(request, 'permits', 'view', api=True)
    if redir:
        return redir
    user = _get_user(request)
    if not user:
        return JsonResponse({'error': 'Login required'}, status=401)
    role = user.get('role', '')

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
        except Exception:
            return JsonResponse({'error': 'Invalid JSON'}, status=400)

        action = data.get('action')

        # ── CREATE ──────────────────────────────────────────────────────
        if action == 'create':
            if role not in ('maintenance_engineer', 'admin'):
                return JsonResponse({'error': 'Maintenance engineer role required'}, status=403)
            permit = create_permit({
                'receiver':               user['username'],
                'receiver_name':          user['name'],
                'receiver_company':       data.get('receiver_company', 'POWERCHINA'),
                'receiver_id':            data.get('receiver_id', ''),
                'job_description':        data.get('job_description', ''),
                'location':               data.get('location', ''),
                'equipment':              data.get('equipment', ''),
                'work_type':              data.get('work_type', 'electrical'),
                'tools_equipment':        data.get('tools_equipment', ''),
                'expected_duration':      data.get('expected_duration', ''),
                'num_employees':          data.get('num_employees', ''),
                'sld_drawing_no':         data.get('sld_drawing_no', ''),
                'energized_lines':        data.get('energized_lines', False),
                'de_energized_lines':     data.get('de_energized_lines', True),
                'risks':                  data.get('risks', {}),
                'docs_to_attach':         data.get('docs_to_attach', {}),
                'precaution_checks':      data.get('precaution_checks', {}),
                'inspected_areas':        data.get('inspected_areas', {}),
                'ppe_required':           data.get('ppe_required', {}),
                'hazards':                data.get('hazards', ''),
                'precautions':            data.get('precautions', ''),
                'additional_precautions': data.get('additional_precautions', []),
                'isolation_required':     data.get('isolation_required', False),
                'isolation_details':      data.get('isolation_details', ''),
                'isolation_type':         data.get('isolation_type', {}),
                'isolation_sequence':     data.get('isolation_sequence', []),
                'valid_from':             data.get('valid_from', ''),
                'valid_until':            data.get('valid_until', ''),
                'application_datetime':   data.get('application_datetime', ''),
                'workers':                data.get('workers', ''),
                'workers_list':           data.get('workers_list', []),
                'receiver_signature':     data.get('receiver_signature'),
            })
            # Notify operation engineers (issuers) + HSE for awareness
            op_engineers = get_users_by_role('operation_engineer')
            hse_officers = get_users_by_role('hse_engineer')
            notify_permit_created(permit, op_engineers, hse_officers)
            return JsonResponse({'ok': True, 'permit_id': permit['id']})

        # ── ISSUE (Operation Engineer) ───────────────────────────────────
        elif action == 'issue':
            if role not in ('operation_engineer', 'admin'):
                return JsonResponse({'error': 'Operation engineer role required'}, status=403)
            permit = get_permit(data.get('permit_id', ''))
            if not permit:
                return JsonResponse({'error': 'Permit not found'}, status=404)
            if permit['status'] != 'pending_issue':
                return JsonResponse({'error': 'Permit is not in pending_issue state'}, status=400)

            updates = {
                'status':           'pending_hse',
                'issuer':           user['username'],
                'issuer_name':      user['name'],
                'issuer_signature': data.get('issuer_signature'),
                'issued_at':        datetime.now().isoformat(),
            }
            updated = update_permit(permit['id'], updates)
            # Notify HSE
            hse_list = get_users_by_role('hse_engineer')
            notify_permit_issued(updated, hse_list)
            return JsonResponse({'ok': True, 'permit': updated})

        # ── HSE SIGN-OFF ─────────────────────────────────────────────────
        elif action == 'hse_sign':
            if role not in ('hse_engineer', 'admin'):
                return JsonResponse({'error': 'HSE engineer role required'}, status=403)
            permit = get_permit(data.get('permit_id', ''))
            if not permit:
                return JsonResponse({'error': 'Permit not found'}, status=404)
            if permit['status'] != 'pending_hse':
                return JsonResponse({'error': 'Permit is not awaiting HSE sign-off'}, status=400)

            permit_number = data.get('permit_number') or get_next_permit_number()
            isolation_number = data.get('isolation_cert_number') or (
                get_next_isolation_number() if permit.get('isolation_required') else None
            )
            updates = {
                'status':                'waiting_for_close',
                'permit_number':         permit_number,
                'isolation_cert_number': isolation_number,
                'hse_officer':           user['username'],
                'hse_name':              user['name'],
                'hse_signature':         data.get('hse_signature'),
                'hse_signed_at':         datetime.now().isoformat(),
            }
            updated = update_permit(permit['id'], updates)

            # Notify all duty maintenance engineers
            duty_engineers = _get_duty_engineers()
            notify_permit_approved(updated, duty_engineers)
            return JsonResponse({'ok': True, 'permit': updated})

        # ── CLOSE ────────────────────────────────────────────────────────
        elif action == 'close':
            if role not in ('operation_engineer', 'admin', 'hse_engineer'):
                return JsonResponse({'error': 'Insufficient role'}, status=403)
            permit = get_permit(data.get('permit_id', ''))
            if not permit:
                return JsonResponse({'error': 'Permit not found'}, status=404)
            if permit['status'] not in ('active', 'waiting_for_close'):
                return JsonResponse({'error': 'Permit is not in a closeable state'}, status=400)

            # For waiting_for_close permits, require all uploads
            if permit['status'] == 'waiting_for_close':
                activity_images = data.get('activity_images', [])
                if not activity_images:
                    return JsonResponse({'error': 'Activity images are required before closing'}, status=400)
                # Work Started signatures must exist from earlier workflow steps
                if not permit.get('receiver_signature'):
                    return JsonResponse({'error': 'Receiver work-started signature is missing'}, status=400)
                if not permit.get('issuer_signature'):
                    return JsonResponse({'error': 'Issuer work-started signature is missing'}, status=400)
                if not permit.get('hse_signature'):
                    return JsonResponse({'error': 'HSE work-started signature is missing'}, status=400)
                # Closure signatures must be uploaded now
                closure_rcv = data.get('closure_receiver_signature')
                closure_iss = data.get('closure_issuer_signature')
                closure_hse = data.get('closure_hse_signature')
                if not closure_rcv:
                    return JsonResponse({'error': 'Closure receiver signature is required'}, status=400)
                if not closure_iss:
                    return JsonResponse({'error': 'Closure issuer signature is required'}, status=400)
                if not closure_hse:
                    return JsonResponse({'error': 'Closure HSE signature is required'}, status=400)
                updated = update_permit(permit['id'], {
                    'status': 'closed',
                    'closed_at':      datetime.now().isoformat(),
                    'closed_by':      user['username'],
                    'closed_by_name': user['name'],
                    'activity_images':            activity_images,
                    'closure_receiver_signature': closure_rcv,
                    'closure_issuer_signature':   closure_iss,
                    'closure_hse_signature':      closure_hse,
                })
            else:
                updated = update_permit(permit['id'], {
                    'status': 'closed',
                    'closed_at': datetime.now().isoformat(),
                })
            return JsonResponse({'ok': True, 'permit': updated})

        # ── CANCEL ───────────────────────────────────────────────────────
        elif action == 'cancel':
            permit = get_permit(data.get('permit_id', ''))
            if not permit:
                return JsonResponse({'error': 'Permit not found'}, status=404)
            if permit.get('receiver') != user['username'] and role not in ('admin',):
                return JsonResponse({'error': 'Not authorized'}, status=403)
            if permit['status'] not in ('pending_issue',):
                return JsonResponse({'error': 'Can only cancel pending permits'}, status=400)
            updated = update_permit(permit['id'], {'status': 'cancelled'})
            return JsonResponse({'ok': True, 'permit': updated})

    return JsonResponse({'error': 'Bad request'}, status=400)


def cmms_permit_pdf(request, permit_id):
    """Export permit as PDF."""
    redir = _require_login(request)
    if redir:
        return redir
    redir = _require_module(request, 'permits', 'view')
    if redir:
        return redir

    buf = generate_permit_pdf(permit_id)
    if not buf:
        return HttpResponse('PDF generation failed or reportlab not installed.', status=500)

    permit = get_permit(permit_id)
    pnum = permit.get('permit_number') or permit_id[:8]
    filename = f"permit_{pnum}.pdf".replace(' ', '_')

    resp = HttpResponse(buf.read(), content_type='application/pdf')
    resp['Content-Disposition'] = f'inline; filename="{filename}"'
    return resp


# ── EMAIL CONFIG API (admin only) ─────────────────────────────────────────

@csrf_exempt
def cmms_email_config_api(request):
    """Admin: update email settings."""
    user = _get_user(request)
    if not user or user.get('role') != 'admin':
        return JsonResponse({'error': 'Admin only'}, status=403)

    config_file = Path(settings.BASE_DIR) / 'cmms_email_config.json'

    if request.method == 'GET':
        if config_file.exists():
            with open(config_file) as f:
                cfg = json.load(f)
            # Mask password
            cfg['password'] = '••••••••' if cfg.get('password') else ''
            return JsonResponse({'ok': True, 'config': cfg})
        return JsonResponse({'ok': True, 'config': {}})

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
        except Exception:
            return JsonResponse({'error': 'Invalid JSON'}, status=400)
        # Save config (password only updated if not masked)
        existing = {}
        if config_file.exists():
            with open(config_file) as f:
                existing = json.load(f)
        cfg = {
            'host':     data.get('host', existing.get('host', 'smtp.gmail.com')),
            'port':     data.get('port', existing.get('port', 587)),
            'use_tls':  data.get('use_tls', existing.get('use_tls', True)),
            'username': data.get('username', existing.get('username', '')),
            'password': data.get('password', existing.get('password', ''))
                        if data.get('password') and data['password'] != '••••••••'
                        else existing.get('password', ''),
            'from_email': data.get('from_email', existing.get('from_email', '')),
        }
        with open(config_file, 'w') as f:
            json.dump(cfg, f, indent=2)
        return JsonResponse({'ok': True})

    return JsonResponse({'error': 'Bad request'}, status=400)


# ── HELPERS ───────────────────────────────────────────────────────────────

def _get_duty_engineers():
    """
    Return list of maintenance engineers on duty today
    based on schedule_store.json (Day or Night shift).
    Falls back to all maintenance engineers if schedule unavailable.
    """
    from pathlib import Path
    import json as _json

    schedule_file = Path(settings.BASE_DIR) / 'schedule_store.json'
    today = datetime.now().strftime('%Y-%m-%d')

    try:
        with open(schedule_file, 'r') as f:
            store = _json.load(f)
        engineers_on_duty = []
        for eng in store.get('engineers', []):
            shifts = eng.get('schedule', {})
            shift_today = shifts.get(today, '')
            if shift_today and shift_today.upper() not in ('R', 'REST', '', 'OFF'):
                # Look up by name
                all_users = get_all_users()
                matched = next(
                    (u for u in all_users
                     if u['name'].strip().lower() == eng.get('name', '').strip().lower()
                     and u['role'] == 'maintenance_engineer'),
                    None
                )
                if matched:
                    engineers_on_duty.append(matched)
        if engineers_on_duty:
            return engineers_on_duty
    except Exception:
        pass

    # Fallback: all maintenance engineers
    return get_users_by_role('maintenance_engineer')


# ── MANPOWER DUTY API ─────────────────────────────────────────────────────

def cmms_duty_staff_api(request):
    """Return engineers and technicians on duty for a given date from schedule."""
    redir = _require_module(request, any_of=('activities', 'permits', 'handover'), api=True)
    if redir:
        return redir
    date_str = request.GET.get('date', datetime.now().strftime('%Y-%m-%d'))
    staff = get_duty_staff(date_str)
    return JsonResponse({'ok': True, 'date': date_str, **staff})


# ── ICC PDF DOWNLOAD ──────────────────────────────────────────────────────

def cmms_icc_pdf(request, permit_id):
    """Download the Isolation Confirmation Certificate PDF."""
    redir = _require_login(request)
    if redir:
        return redir
    redir = _require_module(request, 'permits', 'view')
    if redir:
        return redir
    buf = generate_icc_pdf(permit_id)
    if not buf:
        return HttpResponse('ICC PDF failed or isolation not required.', status=400)
    permit = get_permit(permit_id)
    icc_no = permit.get('isolation_cert_number') or permit_id[:8]
    filename = f"ICC_{icc_no}.pdf".replace(' ', '_')
    resp = HttpResponse(buf.read(), content_type='application/pdf')
    resp['Content-Disposition'] = f'inline; filename="{filename}"'
    return resp


# ── PERMIT WORD EXPORT ────────────────────────────────────────────────────

def cmms_permit_docx(request, permit_id):
    """Download the filled MP-10 General Work Permit as a Word (.docx) file."""
    redir = _require_login(request)
    if redir:
        return redir
    redir = _require_module(request, 'permits', 'view')
    if redir:
        return redir
    try:
        buf = generate_permit_docx(permit_id)
    except FileNotFoundError as e:
        return HttpResponse(str(e), status=404)
    except ValueError as e:
        return HttpResponse(str(e), status=404)
    except Exception as e:
        return HttpResponse(f'Error generating Word document: {e}', status=500)
    permit = get_permit(permit_id)
    permit_no = (permit or {}).get('permit_no') or permit_id[:8].upper()
    filename = f"Permit_{permit_no}.docx".replace(' ', '_')
    resp = HttpResponse(
        buf.read(),
        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
    resp['Content-Disposition'] = f'attachment; filename="{filename}"'
    return resp


# ── ACTIVITY EMAIL TRIGGER ────────────────────────────────────────────────

@csrf_exempt
def cmms_send_activity_email(request, record_id):
    """
    Manually or auto-trigger email to engineer for today's activity.
    Called from the activity detail page on first load for today.
    """
    redir = _require_module(request, 'activities', 'view', api=True)
    if redir:
        return redir
    user = _get_user(request)

    record = get_record(record_id)
    if not record:
        return JsonResponse({'error': 'Record not found'}, status=404)

    if record.get('email_sent'):
        return JsonResponse({'ok': True, 'skipped': True})

    activity = get_activity(record['activity_id'])
    if not activity:
        return JsonResponse({'error': 'Activity not found'}, status=404)

    # Find engineer's email
    from .auth_utils import get_user_detail
    from .email_utils import notify_activity_assigned
    eng = get_user_detail(record.get('engineer', ''))
    if eng and eng.get('email'):
        notify_activity_assigned(activity, eng['email'], record.get('date'))

    update_record(record_id, {'email_sent': True})
    return JsonResponse({'ok': True, 'sent': True})


# ── HANDOVER / SHIFT LOG ──────────────────────────────────────────────────

def cmms_handovers(request):
    """List all handover reports, grouped by date (newest first)."""
    redir = _require_login(request)
    if redir:
        return redir
    redir = _require_module(request, 'handover', 'view')
    if redir:
        return redir
    user = _get_user(request)

    dates = get_handover_dates()
    # Build day-wise pairs
    date_groups = []
    for d in dates:
        pair = get_handovers_by_date(d)
        from datetime import date as _date, datetime as _dt
        try:
            d_obj = _date.fromisoformat(d)
            label = d_obj.strftime('%A, %d %B %Y')
        except Exception:
            label = d
        date_groups.append({'date': d, 'label': label,
                             'day': pair['day'], 'night': pair['night']})

    ctx = _ctx(request, {
        'date_groups': date_groups,
        'today': datetime.now().strftime('%Y-%m-%d'),
    })
    return render(request, 'core/cmms_handovers.html', ctx)


def cmms_handover_detail(request, handover_id=None):
    """Create new or view/edit existing handover report."""
    redir = _require_login(request)
    if redir:
        return redir
    redir = _require_module(request, 'handover', 'view')
    if redir:
        return redir
    user = _get_user(request)

    handover = get_handover(handover_id) if handover_id else None

    today = datetime.now().strftime('%Y-%m-%d')
    # Duty staff for dropdowns
    duty = get_duty_staff(today)
    from .cmms_utils import get_all_technicians_from_schedule
    all_techs = get_all_technicians_from_schedule(today)
    # All engineers from schedule_store
    from pathlib import Path as _P
    import json as _j
    _sf = _P(settings.BASE_DIR) / 'schedule_store.json'
    all_engineers = []
    if _sf.exists():
        _store = _j.loads(_sf.read_text(encoding='utf-8'))
        all_engineers = [e.get('name', '') for e in _store.get('engineers', [])]

    ctx = _ctx(request, {
        'handover': handover,
        'today': today,
        'all_engineers': all_engineers,
        'all_techs': all_techs,
        'duty_engineers': duty.get('engineers', []),
        'duty_technicians': duty.get('technicians', []),
    })
    return render(request, 'core/cmms_handover_detail.html', ctx)


@csrf_exempt
def cmms_handover_api(request):
    """Create / update handover records."""
    redir = _require_login(request)
    if redir:
        return JsonResponse({'error': 'Login required'}, status=401)
    redir = _require_module(request, 'handover', 'view', api=True)
    if redir:
        return redir
    user = _get_user(request)

    if request.method == 'POST':
        # Multipart (image upload) handled separately
        try:
            data = json.loads(request.body)
        except Exception:
            return JsonResponse({'error': 'Invalid JSON'}, status=400)

        action = data.get('action', 'create')

        if action == 'create':
            # Prevent duplicate (same date+shift)
            existing = get_handovers_by_date(data.get('date', ''))
            shift = data.get('shift', 'day')
            if existing.get(shift):
                # Update instead of duplicate
                updated = update_handover(existing[shift]['id'], {
                    k: v for k, v in data.items() if k not in ('action',)
                })
                return JsonResponse({'ok': True, 'handover_id': updated['id'], 'action': 'updated'})
            data['created_by'] = user['username']
            h = create_handover(data)
            return JsonResponse({'ok': True, 'handover_id': h['id'], 'action': 'created'})

        elif action == 'update':
            hid = data.get('handover_id', '')
            updates = {k: v for k, v in data.items() if k not in ('action', 'handover_id')}
            h = update_handover(hid, updates)
            if not h:
                return JsonResponse({'error': 'Not found'}, status=404)
            return JsonResponse({'ok': True, 'handover_id': h['id']})

        elif action == 'submit':
            hid = data.get('handover_id', '')
            h = update_handover(hid, {
                'status': 'submitted',
                'submitted_at': datetime.now().isoformat(),
                'shift_engineer_sig': data.get('shift_engineer_sig', ''),
                'incoming_engineer_sig': data.get('incoming_engineer_sig', ''),
            })
            if not h:
                return JsonResponse({'error': 'Not found'}, status=404)
            return JsonResponse({'ok': True})

        elif action == 'delete':
            hid = data.get('handover_id', '')
            from .cmms_utils import HANDOVERS_FILE, _save
            all_h = get_handovers()
            all_h = [h for h in all_h if h['id'] != hid]
            _save(HANDOVERS_FILE, all_h)
            return JsonResponse({'ok': True})

    return JsonResponse({'error': 'Bad request'}, status=400)


@csrf_exempt
def cmms_handover_image_api(request, handover_id):
    """Upload observation images for a handover."""
    redir = _require_login(request)
    if redir:
        return JsonResponse({'error': 'Login required'}, status=401)
    redir = _require_module(request, 'handover', 'view', api=True)
    if redir:
        return redir

    if request.method == 'POST':
        imgs = []
        for key in request.FILES:
            f = request.FILES[key]
            rel = save_handover_image(handover_id, f)
            imgs.append(rel)
        return JsonResponse({'ok': True, 'images': imgs})

    if request.method == 'DELETE':
        import urllib.parse
        rel = request.GET.get('img', '')
        h = get_handover(handover_id)
        if h and rel:
            imgs = [i for i in h.get('observation_images', []) if i != rel]
            update_handover(handover_id, {'observation_images': imgs})
            # Delete file
            from .cmms_utils import MEDIA_ROOT as _MR
            fpath = _MR / rel
            if fpath.exists():
                fpath.unlink()
        return JsonResponse({'ok': True})

    return JsonResponse({'error': 'Method not allowed'}, status=405)
