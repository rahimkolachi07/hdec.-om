"""
project_views/cmms/permits.py — Permit views for a project.
"""
import json
from datetime import datetime
from django.shortcuts import render, redirect
from django.http import JsonResponse, HttpResponse
from django.views.decorators.csrf import csrf_exempt

from ..base import _pctx, _guard, WORK_TYPES, PERMIT_STATUSES
from ...views import get_user, login_required
from ...project_data import (
    permit_load, permit_get, permit_create, permit_update,
    permit_delete, permit_export_excel,
)


@login_required
def project_permits(request, country_id, project_id, category='maintenance'):
    country, project = _guard(request, country_id, project_id, 'permits', 'view')
    if not country:
        return redirect('/')

    status_filter = request.GET.get('status', '')
    permits = permit_load(country_id, project_id)

    if status_filter:
        permits = [p for p in permits if p.get('status') == status_filter]

    for p in permits:
        p['status_label'] = PERMIT_STATUSES.get(p.get('status', ''), p.get('status', ''))
        p['work_type_label'] = WORK_TYPES.get(p.get('work_type', ''), p.get('work_type', ''))

    user = get_user(request)
    role = user.get('role', '') if user else ''
    can_create = role in ('admin', 'maintenance_engineer')

    ctx, _, _ = _pctx(request, country_id, project_id, {
        'permits': permits,
        'permit_statuses': PERMIT_STATUSES,
        'status_filter': status_filter,
        'can_create': can_create,
    }, category=category)
    return render(request, 'core/modules/cmms/permits/list.html', ctx)


@login_required
def project_permit_detail(request, country_id, project_id, permit_id=None, category='maintenance'):
    country, project = _guard(request, country_id, project_id, 'permits', 'view')
    if not country:
        return redirect('/')

    user = get_user(request)
    role = user.get('role', '') if user else ''

    permit = None
    if permit_id and permit_id != 'new':
        permit = permit_get(country_id, project_id, permit_id)
        if not permit:
            return redirect(f'/p/{country_id}/{project_id}/{category}/cmms/permits/')

    ctx, _, _ = _pctx(request, country_id, project_id, {
        'permit': permit,
        'work_types': WORK_TYPES,
        'permit_statuses': PERMIT_STATUSES,
        'can_issue': role in ('admin', 'operation_engineer'),
        'can_hse': role in ('admin', 'hse_engineer'),
        'can_close': role in ('admin', 'maintenance_engineer', 'hse_engineer', 'operation_engineer'),
        'can_create': role in ('admin', 'maintenance_engineer'),
        'is_new': permit is None,
    }, category=category)
    return render(request, 'core/modules/cmms/permits/detail.html', ctx)


@csrf_exempt
@login_required
def project_permits_api(request, country_id, project_id, category='maintenance'):
    level = 'view' if request.method == 'GET' else 'edit'
    country, project = _guard(request, country_id, project_id, 'permits', level)
    if not country:
        return JsonResponse({'error': 'Forbidden'}, status=403)

    user = get_user(request)
    role = user.get('role', '') if user else ''

    if request.method == 'GET':
        action = request.GET.get('action', '')
        if action == 'export':
            xls = permit_export_excel(country_id, project_id)
            resp = HttpResponse(xls, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            resp['Content-Disposition'] = f'attachment; filename="{project_id}_permits.xlsx"'
            return resp
        return JsonResponse({'permits': permit_load(country_id, project_id)})

    if request.method != 'POST':
        return JsonResponse({'error': 'Method not allowed'}, status=405)

    try:
        data = json.loads(request.body)
    except Exception:
        return JsonResponse({'error': 'Invalid JSON'}, status=400)

    action = data.get('action', '')
    permit_id = data.get('permit_id', '')

    if action == 'create':
        if role not in ('admin', 'maintenance_engineer'):
            return JsonResponse({'error': 'Forbidden'}, status=403)
        data['receiver'] = user.get('username', '')
        data['receiver_name'] = user.get('name', '')
        p = permit_create(country_id, project_id, data)
        return JsonResponse({'ok': True, 'permit': p})

    elif action == 'issue':
        if role not in ('admin', 'operation_engineer'):
            return JsonResponse({'error': 'Forbidden'}, status=403)
        p = permit_get(country_id, project_id, permit_id)
        if not p or p['status'] != 'pending_issue':
            return JsonResponse({'error': 'Invalid permit state'})
        updated = permit_update(country_id, project_id, permit_id, {
            'status': 'pending_hse',
            'issuer': user.get('username', ''),
            'issuer_name': user.get('name', ''),
            'permit_number': data.get('permit_number', ''),
            'isolation_cert_number': data.get('isolation_cert_number', ''),
            'valid_from': data.get('valid_from', ''),
            'valid_until': data.get('valid_until', ''),
            'issuer_signature': data.get('issuer_signature', ''),
            'issued_at': datetime.now().isoformat(),
        })
        return JsonResponse({'ok': True, 'permit': updated})

    elif action == 'hse_sign':
        if role not in ('admin', 'hse_engineer'):
            return JsonResponse({'error': 'Forbidden'}, status=403)
        p = permit_get(country_id, project_id, permit_id)
        if not p or p['status'] != 'pending_hse':
            return JsonResponse({'error': 'Invalid permit state'})
        updated = permit_update(country_id, project_id, permit_id, {
            'status': 'waiting_for_close',
            'hse_officer': user.get('username', ''),
            'hse_name': user.get('name', ''),
            'hse_signature': data.get('hse_signature', ''),
            'hse_signed_at': datetime.now().isoformat(),
        })
        return JsonResponse({'ok': True, 'permit': updated})

    elif action == 'close':
        p = permit_get(country_id, project_id, permit_id)
        if not p or p['status'] not in ('active', 'waiting_for_close'):
            return JsonResponse({'error': 'Invalid permit state'})
        updated = permit_update(country_id, project_id, permit_id, {
            'status': 'closed',
            'closed_at': datetime.now().isoformat(),
            'closed_by': user.get('username', ''),
            'closed_by_name': user.get('name', ''),
            'closure_receiver_signature': data.get('closure_receiver_signature', ''),
            'closure_issuer_signature': data.get('closure_issuer_signature', ''),
            'closure_hse_signature': data.get('closure_hse_signature', ''),
            'activity_images': data.get('activity_images', []),
        })
        return JsonResponse({'ok': True, 'permit': updated})

    elif action == 'cancel':
        if role not in ('admin', 'operation_engineer'):
            return JsonResponse({'error': 'Forbidden'}, status=403)
        updated = permit_update(country_id, project_id, permit_id, {'status': 'cancelled'})
        return JsonResponse({'ok': True, 'permit': updated})

    return JsonResponse({'error': f'Unknown action: {action}'}, status=400)
