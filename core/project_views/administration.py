"""
Project-scoped Administration module views.
"""
import json

from django.http import JsonResponse
from django.shortcuts import redirect, render
from django.views.decorators.csrf import csrf_exempt

from ..admin_modules_data import (
    list_vehicles, get_vehicle, create_vehicle, update_vehicle, delete_vehicle,
    list_residences, get_residence, create_residence, update_residence, delete_residence,
    list_workforce, get_workforce, create_workforce, update_workforce, delete_workforce,
    list_gatepasses, get_gatepass, create_gatepass, update_gatepass, delete_gatepass,
    list_equipment, get_equipment, create_equipment, update_equipment, delete_equipment,
    list_trainings, get_training, create_training, update_training, delete_training,
)
from ..auth_utils import DEFAULT_PERMISSIONS, MODULES, has_permission
from ..project_utils import MODULE_META, get_category_modules
from ..views import get_user, login_required
from .base import _guard, _pctx

ADMIN_MODULE_HANDLERS = {
    'vehicles': (list_vehicles, get_vehicle, create_vehicle, update_vehicle, delete_vehicle),
    'residences': (list_residences, get_residence, create_residence, update_residence, delete_residence),
    'workforce': (list_workforce, get_workforce, create_workforce, update_workforce, delete_workforce),
    'gatepasses': (list_gatepasses, get_gatepass, create_gatepass, update_gatepass, delete_gatepass),
    'equipment': (list_equipment, get_equipment, create_equipment, update_equipment, delete_equipment),
    'trainings': (list_trainings, get_training, create_training, update_training, delete_training),
}


def _admin_module_enabled(project, module_id):
    return module_id in ADMIN_MODULE_HANDLERS and module_id in get_category_modules(project, 'administration')


def _module_access_level(user, module_id):
    if has_permission(user, module_id, 'edit'):
        return 'edit'
    if has_permission(user, module_id, 'view'):
        return 'view'
    return 'none'


@login_required
def project_administration_module(request, country_id, project_id, module_id):
    country, project = _guard(request, country_id, project_id, module_id, 'view')
    if not country:
        return redirect('/')
    if not _admin_module_enabled(project, module_id):
        return redirect(f'/p/{country_id}/{project_id}/administration/')

    user = get_user(request)
    module_meta = MODULE_META[module_id]
    ctx, _, _ = _pctx(request, country_id, project_id, {
        'users_json': '[]',
        'countries_json': '[]',
        'modules_json': json.dumps(MODULES),
        'defaults_json': json.dumps(DEFAULT_PERMISSIONS),
        'admin_template_mode': 'module',
        'admin_api_base': f'/api/p/{country_id}/{project_id}/administration',
        'module_only_mod': module_id,
        'module_meta': module_meta,
        'module_access_level': _module_access_level(user, module_id),
        'module_back_url': f'/p/{country_id}/{project_id}/administration/',
        'module_back_label': 'Administration',
    }, category='administration')
    return render(request, 'core/admin.html', ctx)


@csrf_exempt
@login_required
def project_administration_api(request, country_id, project_id, module_id):
    level = 'view' if request.method == 'GET' else 'edit'
    country, project = _guard(request, country_id, project_id, module_id, level)
    if not country or not _admin_module_enabled(project, module_id):
        return JsonResponse({'error': 'Forbidden'}, status=403)

    list_fn, _, create_fn, _, _ = ADMIN_MODULE_HANDLERS[module_id]
    if request.method == 'GET':
        return JsonResponse({'records': list_fn(country_id, project_id)})

    try:
        data = json.loads(request.body or '{}')
    except Exception:
        return JsonResponse({'error': 'Invalid JSON'}, status=400)

    if request.method == 'POST':
        return JsonResponse({'record': create_fn(data, country_id, project_id)}, status=201)
    return JsonResponse({'error': 'Method not allowed'}, status=405)


@csrf_exempt
@login_required
def project_administration_detail_api(request, country_id, project_id, module_id, rid):
    level = 'view' if request.method == 'GET' else 'edit'
    country, project = _guard(request, country_id, project_id, module_id, level)
    if not country or not _admin_module_enabled(project, module_id):
        return JsonResponse({'error': 'Forbidden'}, status=403)

    _, get_fn, _, update_fn, delete_fn = ADMIN_MODULE_HANDLERS[module_id]

    if request.method == 'GET':
        record = get_fn(rid, country_id, project_id)
        return JsonResponse({'record': record}) if record else JsonResponse({'error': 'Not found'}, status=404)

    try:
        data = json.loads(request.body or '{}')
    except Exception:
        return JsonResponse({'error': 'Invalid JSON'}, status=400)

    if request.method == 'PATCH':
        record = update_fn(rid, data, country_id, project_id)
        return JsonResponse({'record': record}) if record else JsonResponse({'error': 'Not found'}, status=404)
    if request.method == 'DELETE':
        ok = delete_fn(rid, country_id, project_id)
        return JsonResponse({'ok': ok}) if ok else JsonResponse({'error': 'Not found'}, status=404)
    return JsonResponse({'error': 'Method not allowed'}, status=405)
