"""
project_views/handover.py -- Project-scoped shift handover views and API.
"""
from __future__ import annotations

import json
from datetime import datetime

from django.http import Http404, HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.views.decorators.csrf import csrf_exempt

from ..auth_utils import has_permission
from ..project_data import (
    handover_create,
    handover_delete,
    handover_export_excel,
    handover_find_by_date_shift,
    handover_get,
    handover_list,
    handover_update,
)
from ..views import get_user, login_required
from .base import _guard, _pctx


def _display_user(user: dict | None) -> str:
    return (user or {}).get('name') or (user or {}).get('username') or ''


@login_required
def project_cmms_handover_list(request, country_id, project_id, category='maintenance'):
    country, project = _guard(request, country_id, project_id, 'handover', 'view')
    if not country:
        return redirect('/')

    user = get_user(request)
    ctx, _, _ = _pctx(request, country_id, project_id, {
        'handovers': handover_list(country_id, project_id),
        'can_create': has_permission(user, 'handover', 'edit'),
    }, category=category)
    return render(request, 'core/modules/cmms/handover/list.html', ctx)


@login_required
def project_cmms_handover_new(request, country_id, project_id, category='maintenance'):
    country, project = _guard(request, country_id, project_id, 'handover', 'edit')
    if not country:
        return redirect('/')

    ctx, _, _ = _pctx(request, country_id, project_id, {
        'handover': None,
        'handover_technicians_json': '[]',
        'is_new': True,
        'can_edit': True,
        'is_editable': True,
    }, category=category)
    return render(request, 'core/modules/cmms/handover/detail.html', ctx)


@login_required
def project_cmms_handover_detail(request, country_id, project_id, category='maintenance', handover_id=''):
    country, project = _guard(request, country_id, project_id, 'handover', 'view')
    if not country:
        return redirect('/')

    handover = handover_get(country_id, project_id, handover_id)
    if not handover:
        raise Http404('Handover not found')

    user = get_user(request)
    can_edit = has_permission(user, 'handover', 'edit')
    is_editable = can_edit and handover.get('status') == 'draft'
    ctx, _, _ = _pctx(request, country_id, project_id, {
        'handover': handover,
        'handover_technicians_json': json.dumps(handover.get('technicians', [])),
        'is_new': False,
        'can_edit': can_edit,
        'is_editable': is_editable,
    }, category=category)
    return render(request, 'core/modules/cmms/handover/detail.html', ctx)


@csrf_exempt
@login_required
def project_cmms_handover_api(request, country_id, project_id, category='maintenance'):
    level = 'view' if request.method == 'GET' else 'edit'
    country, project = _guard(request, country_id, project_id, 'handover', level)
    if not country:
        return JsonResponse({'error': 'Forbidden'}, status=403)

    if request.method == 'GET':
        action = request.GET.get('action', '').strip().lower()
        if action == 'export':
            xls = handover_export_excel(country_id, project_id)
            response = HttpResponse(
                xls,
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
            response['Content-Disposition'] = f'attachment; filename="{project_id}_handover.xlsx"'
            return response
        return JsonResponse({'handovers': handover_list(country_id, project_id)})

    if request.method != 'POST':
        return JsonResponse({'error': 'Method not allowed'}, status=405)

    try:
        data = json.loads(request.body)
    except Exception:
        return JsonResponse({'error': 'Invalid JSON'}, status=400)

    user = get_user(request)
    actor = _display_user(user)
    action = str(data.get('action', '') or '').strip().lower()
    status = 'submitted' if str(data.get('status', '')).strip().lower() == 'submitted' else 'draft'
    payload = {
        'date': data.get('date', ''),
        'shift': data.get('shift', ''),
        'timing': data.get('timing', ''),
        'shift_incharge': data.get('shift_incharge', ''),
        'technicians': data.get('technicians', []),
        'major_alarms': data.get('major_alarms', ''),
        'equipment_breakdown': data.get('equipment_breakdown', ''),
        'maintenance_activities': data.get('maintenance_activities', ''),
        'inverter_faults': data.get('inverter_faults', ''),
        'scb_faults': data.get('scb_faults', ''),
        'spare_parts': data.get('spare_parts', ''),
        'key_issues': data.get('key_issues', ''),
        'pending_work': data.get('pending_work', ''),
        'instructions_next_shift': data.get('instructions_next_shift', ''),
        'observation_text': data.get('observation_text', ''),
        'status': status,
        'shift_engineer_sig': data.get('shift_engineer_sig', ''),
        'incoming_engineer_sig': data.get('incoming_engineer_sig', ''),
        'submitted_at': data.get('submitted_at', '') or (datetime.now().isoformat() if status == 'submitted' else ''),
    }

    if action == 'create':
        existing = handover_find_by_date_shift(country_id, project_id, payload['date'], payload['shift'])
        if existing and existing.get('status') == 'submitted':
            return JsonResponse({'error': 'A submitted handover already exists for this date and shift.'}, status=400)
        if existing:
            handover = handover_update(country_id, project_id, existing['id'], payload, actor)
        else:
            handover = handover_create(country_id, project_id, payload, actor)
        return JsonResponse({'ok': True, 'handover': handover})

    if action == 'update':
        handover_id = str(data.get('id', '') or '').strip()
        if not handover_id:
            return JsonResponse({'error': 'Handover id is required.'}, status=400)
        current = handover_get(country_id, project_id, handover_id)
        if not current:
            return JsonResponse({'error': 'Handover not found.'}, status=404)
        if current.get('status') == 'submitted':
            return JsonResponse({'error': 'Submitted handovers are locked.'}, status=400)
        handover = handover_update(country_id, project_id, handover_id, payload, actor)
        return JsonResponse({'ok': True, 'handover': handover})

    if action == 'delete':
        handover_id = str(data.get('id', '') or '').strip()
        if not handover_id:
            return JsonResponse({'error': 'Handover id is required.'}, status=400)
        current = handover_get(country_id, project_id, handover_id, include_deleted=True)
        if not current:
            return JsonResponse({'error': 'Handover not found.'}, status=404)
        if current.get('status') == 'submitted':
            return JsonResponse({'error': 'Submitted handovers cannot be deleted.'}, status=400)
        handover_delete(country_id, project_id, handover_id, actor)
        return JsonResponse({'ok': True})

    return JsonResponse({'error': 'Unknown action'}, status=400)
