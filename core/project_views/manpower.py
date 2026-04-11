"""
project_views/manpower.py — Manpower views for a project.
"""
import json
from django.shortcuts import render, redirect
from django.http import JsonResponse, HttpResponse
from django.views.decorators.csrf import csrf_exempt

from ..views import get_user, login_required
from ..project_data import (
    mp_get, mp_add_person, mp_remove_person, mp_bulk_update,
    mp_parse_excel, mp_export_excel, get_blank_template,
)
from .base import _pctx, _guard


@login_required
def project_manpower(request, country_id, project_id, category='maintenance'):
    country, project = _guard(request, country_id, project_id, 'manpower', 'view')
    if not country:
        return redirect('/')

    data = mp_get(country_id, project_id)
    engineers = data.get('engineers', [])
    technicians = data.get('technicians', [])
    all_dates = sorted({
        d for p in engineers + technicians
        for d in p.get('schedule', {}).keys()
    })

    ctx, _, _ = _pctx(request, country_id, project_id, {
        'engineers': engineers,
        'technicians': technicians,
        'all_dates': all_dates,
        'engineers_json': json.dumps(engineers),
        'technicians_json': json.dumps(technicians),
        'updated_at': data.get('updated_at', ''),
    }, category=category)
    return render(request, 'core/modules/manpower/index.html', ctx)


@csrf_exempt
@login_required
def project_manpower_api(request, country_id, project_id, category='maintenance'):
    level = 'view' if request.method == 'GET' else 'edit'
    country, project = _guard(request, country_id, project_id, 'manpower', level)
    if not country:
        return JsonResponse({'error': 'Forbidden'}, status=403)

    if request.method == 'GET':
        action = request.GET.get('action', '')
        if action == 'export':
            xls = mp_export_excel(country_id, project_id)
            resp = HttpResponse(xls, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            resp['Content-Disposition'] = f'attachment; filename="{project_id}_manpower.xlsx"'
            return resp
        if action == 'template':
            xls = get_blank_template('manpower')
            resp = HttpResponse(xls, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            resp['Content-Disposition'] = 'attachment; filename="manpower_template.xlsx"'
            return resp
        return JsonResponse(mp_get(country_id, project_id))

    if request.method == 'POST':
        content_type = request.content_type or ''
        if 'multipart' in content_type or request.FILES.get('file'):
            f = request.FILES.get('file')
            if not f:
                return JsonResponse({'error': 'No file provided'}, status=400)
            try:
                result = mp_parse_excel(f.read())
                mp_bulk_update(country_id, project_id, result['engineers'], result['technicians'])
                return JsonResponse({
                    'ok': True,
                    'engineers': len(result['engineers']),
                    'technicians': len(result['technicians']),
                    'msg': f"Imported {len(result['engineers'])} engineers and {len(result['technicians'])} technicians.",
                })
            except Exception as e:
                return JsonResponse({'error': f'Import failed: {str(e)}'}, status=400)

        try:
            data = json.loads(request.body)
        except Exception:
            return JsonResponse({'error': 'Invalid JSON'}, status=400)

        action = data.get('action', '')
        user = get_user(request)

        if action == 'add_person':
            mp_add_person(country_id, project_id,
                          data.get('category', 'engineers'),
                          data.get('name', ''), data.get('role', ''), data.get('dept', ''))
            return JsonResponse({'ok': True, 'data': mp_get(country_id, project_id)})

        elif action == 'remove_person':
            mp_remove_person(country_id, project_id,
                             data.get('category', 'engineers'), data.get('id', ''))
            return JsonResponse({'ok': True, 'data': mp_get(country_id, project_id)})

        elif action == 'save_all':
            mp_bulk_update(country_id, project_id, data.get('engineers', []), data.get('technicians', []))
            return JsonResponse({'ok': True})

    return JsonResponse({'error': 'Method not allowed'}, status=405)
