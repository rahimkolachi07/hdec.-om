"""
project_views/store.py - Store module views for a project.
"""
import json

from django.http import JsonResponse
from django.shortcuts import redirect, render
from django.views.decorators.csrf import csrf_exempt

from ..views import login_required
from ..project_data import store_create, store_delete, store_get, store_load, store_update
from .base import _guard, _pctx


@login_required
def project_store(request, country_id, project_id, category='maintenance'):
    country, project = _guard(request, country_id, project_id, 'store', 'view')
    if not country:
        return redirect('/')

    items = store_load(country_id, project_id)
    ctx, _, _ = _pctx(request, country_id, project_id, {
        'store_items': items,
        'store_items_json': json.dumps(items),
    }, category=category)
    return render(request, 'core/modules/store/index.html', ctx)


@csrf_exempt
@login_required
def project_store_api(request, country_id, project_id, category='maintenance'):
    level = 'view' if request.method == 'GET' else 'edit'
    country, project = _guard(request, country_id, project_id, 'store', level)
    if not country:
        return JsonResponse({'error': 'Forbidden'}, status=403)

    if request.method == 'GET':
        return JsonResponse({'items': store_load(country_id, project_id)})

    if request.method != 'POST':
        return JsonResponse({'error': 'Method not allowed'}, status=405)

    content_type = request.content_type or ''
    if 'multipart' in content_type or request.FILES:
        action = request.POST.get('action', '')
        pictures = request.FILES.getlist('pictures')
        if action == 'create':
            item = store_create(country_id, project_id, request.POST, pictures)
            return JsonResponse({'ok': True, 'item': item})
        if action == 'update':
            payload = {
                'equipment_name': request.POST.get('equipment_name', ''),
                'date': request.POST.get('date', ''),
                'details': request.POST.get('details', ''),
                'quantity': request.POST.get('quantity', ''),
                'status': request.POST.get('status', 'given'),
                'retain_pictures': request.POST.getlist('retain_pictures'),
            }
            item = store_update(country_id, project_id, request.POST.get('id', ''), payload, pictures)
            if not item:
                return JsonResponse({'error': 'Item not found'}, status=404)
            return JsonResponse({'ok': True, 'item': item})
        return JsonResponse({'error': f'Unknown action: {action}'}, status=400)

    try:
        data = json.loads(request.body)
    except Exception:
        return JsonResponse({'error': 'Invalid JSON'}, status=400)

    action = data.get('action', '')
    if action == 'delete':
        store_delete(country_id, project_id, data.get('id', ''))
        return JsonResponse({'ok': True})
    if action == 'get':
        item = store_get(country_id, project_id, data.get('id', ''))
        if not item:
            return JsonResponse({'error': 'Item not found'}, status=404)
        return JsonResponse({'ok': True, 'item': item})
    return JsonResponse({'error': f'Unknown action: {action}'}, status=400)
