"""
project_views/cmms/handover.py — Shift handover views for a project.
"""
import json
from datetime import datetime
from django.shortcuts import render, redirect
from django.http import JsonResponse, HttpResponse
from django.views.decorators.csrf import csrf_exempt

from ..base import _pctx, _guard
from ...views import get_user, login_required
from ...project_data import (
    ho_load, ho_get, ho_create, ho_update, ho_delete,
    ho_save_one, ho_export_excel,
)


@login_required
def project_handovers(request, country_id, project_id, category='maintenance'):
    country, project = _guard(request, country_id, project_id, 'handover', 'view')
    if not country:
        return redirect('/')

    entries = ho_load(country_id, project_id)
    user = get_user(request)
    role = user.get('role', '') if user else ''

    ctx, _, _ = _pctx(request, country_id, project_id, {
        'handovers': entries,
        'can_create': role in ('admin', 'maintenance_engineer', 'operation_engineer', 'hse_engineer'),
    }, category=category)
    return render(request, 'core/modules/cmms/handover/list.html', ctx)


@login_required
def project_handover_detail(request, country_id, project_id, handover_id=None, category='maintenance'):
    country, project = _guard(request, country_id, project_id, 'handover', 'view')
    if not country:
        return redirect('/')

    handover = None
    if handover_id and handover_id != 'new':
        handover = ho_get(country_id, project_id, handover_id)
        if not handover:
            return redirect(f'/p/{country_id}/{project_id}/{category}/cmms/handover/')

    ctx, _, _ = _pctx(request, country_id, project_id, {
        'handover': handover,
        'is_new': handover is None,
    }, category=category)
    return render(request, 'core/modules/cmms/handover/detail.html', ctx)


@csrf_exempt
@login_required
def project_handover_api(request, country_id, project_id, category='maintenance'):
    level = 'view' if request.method == 'GET' else 'edit'
    country, project = _guard(request, country_id, project_id, 'handover', level)
    if not country:
        return JsonResponse({'error': 'Forbidden'}, status=403)

    user = get_user(request)

    if request.method == 'GET':
        action = request.GET.get('action', '')
        if action == 'export':
            xls = ho_export_excel(country_id, project_id)
            resp = HttpResponse(xls, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            resp['Content-Disposition'] = f'attachment; filename="{project_id}_handover.xlsx"'
            return resp
        return JsonResponse({'handovers': ho_load(country_id, project_id)})

    if request.method != 'POST':
        return JsonResponse({'error': 'Method not allowed'}, status=405)

    try:
        data = json.loads(request.body)
    except Exception:
        return JsonResponse({'error': 'Invalid JSON'}, status=400)

    action = data.get('action', '')
    ho_id = data.get('id', '')

    if action == 'create':
        data['created_by'] = user.get('name', '') if user else ''
        h = ho_create(country_id, project_id, data)
        return JsonResponse({'ok': True, 'handover': h})

    elif action == 'update':
        fields = {k: v for k, v in data.items() if k not in ('action', 'id')}
        updated = ho_update(country_id, project_id, ho_id, fields)
        return JsonResponse({'ok': True, 'handover': updated})

    elif action == 'submit':
        updated = ho_update(country_id, project_id, ho_id, {
            'status': 'submitted',
            'submitted_at': datetime.now().isoformat(),
            'shift_engineer_sig': data.get('shift_engineer_sig', ''),
        })
        return JsonResponse({'ok': True, 'handover': updated})

    elif action == 'delete':
        ho_delete(country_id, project_id, ho_id)
        return JsonResponse({'ok': True})

    return JsonResponse({'error': f'Unknown action: {action}'}, status=400)
