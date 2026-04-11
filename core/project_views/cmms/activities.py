"""
project_views/cmms/activities.py — Activities views for a project.
"""
import json
from django.shortcuts import render, redirect
from django.http import JsonResponse, HttpResponse
from django.views.decorators.csrf import csrf_exempt

from ..base import _pctx, _guard
from ...views import get_user, login_required
from ...project_data import (
    act_load, act_create, act_update, act_delete, act_get,
    act_parse_excel, act_export_excel, get_blank_template,
)


@login_required
def project_activities(request, country_id, project_id, category='maintenance'):
    country, project = _guard(request, country_id, project_id, 'activities', 'view')
    if not country:
        return redirect('/')

    activities = act_load(country_id, project_id)
    ctx, _, _ = _pctx(request, country_id, project_id, {
        'activities': activities,
        'activities_json': json.dumps(activities),
    }, category=category)
    return render(request, 'core/modules/cmms/activities.html', ctx)


@csrf_exempt
@login_required
def project_activities_api(request, country_id, project_id, category='maintenance'):
    level = 'view' if request.method == 'GET' else 'edit'
    country, project = _guard(request, country_id, project_id, 'activities', level)
    if not country:
        return JsonResponse({'error': 'Forbidden'}, status=403)

    user = get_user(request)

    if request.method == 'GET':
        action = request.GET.get('action', '')
        if action == 'export':
            xls = act_export_excel(country_id, project_id)
            resp = HttpResponse(xls, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            resp['Content-Disposition'] = f'attachment; filename="{project_id}_activities.xlsx"'
            return resp
        if action == 'template':
            xls = get_blank_template('activities')
            resp = HttpResponse(xls, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            resp['Content-Disposition'] = 'attachment; filename="activities_template.xlsx"'
            return resp
        return JsonResponse({'activities': act_load(country_id, project_id)})

    if request.method == 'POST':
        content_type = request.content_type or ''
        if 'multipart' in content_type or request.FILES.get('file'):
            f = request.FILES.get('file')
            if not f:
                return JsonResponse({'error': 'No file'}, status=400)
            try:
                parsed = act_parse_excel(f.read())
                for item in parsed:
                    item['created_by'] = user.get('name', '') if user else ''
                    act_create(country_id, project_id, item)
                return JsonResponse({'ok': True, 'imported': len(parsed), 'msg': f'Imported {len(parsed)} activities.'})
            except Exception as e:
                return JsonResponse({'error': f'Import failed: {str(e)}'}, status=400)

        try:
            data = json.loads(request.body)
        except Exception:
            return JsonResponse({'error': 'Invalid JSON'}, status=400)

        action = data.get('action', '')

        if action == 'create':
            data['created_by'] = user.get('name', '') if user else ''
            item = act_create(country_id, project_id, data)
            return JsonResponse({'ok': True, 'activity': item})

        elif action == 'update':
            act_id = data.get('id', '')
            act_update(country_id, project_id, act_id, {k: v for k, v in data.items() if k not in ('action', 'id')})
            return JsonResponse({'ok': True, 'activity': act_get(country_id, project_id, act_id)})

        elif action == 'delete':
            act_delete(country_id, project_id, data.get('id', ''))
            return JsonResponse({'ok': True})

    return JsonResponse({'error': 'Method not allowed'}, status=405)
