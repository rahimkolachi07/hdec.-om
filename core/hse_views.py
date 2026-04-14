"""
HSE Views
API endpoints for HSE module data management.
"""
import json
from django.shortcuts import render
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods

from .views import get_user, login_required
from .auth_utils import has_permission
from .hse_utils import (
    get_permits, get_permit, create_permit, update_permit, delete_permit,
    get_records, get_record, create_record, update_record, delete_record,
)


def _require_hse_permission(request, level='view'):
    user = get_user(request)
    if not user or not has_permission(user, 'sjn_portal', level):
        return JsonResponse({'error': 'Forbidden'}, status=403)
    return None


# ── API Endpoints ─────────────────────────────────────────────────────────

@csrf_exempt
@require_http_methods(["GET", "POST"])
@login_required
def hse_api_permits(request):
    err = _require_hse_permission(request, 'edit' if request.method == 'POST' else 'view')
    if err:
        return err
    if request.method == 'GET':
        permits = get_permits()
        return JsonResponse({'permits': permits})
    elif request.method == 'POST':
        try:
            data = json.loads(request.body)
            permit = create_permit(data)
            return JsonResponse({'permit': permit}, status=201)
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=400)


@csrf_exempt
@require_http_methods(["GET", "PUT", "DELETE"])
@login_required
def hse_api_permit_detail(request, permit_id):
    err = _require_hse_permission(request, 'view' if request.method == 'GET' else 'edit')
    if err:
        return err
    if request.method == 'GET':
        permit = get_permit(permit_id)
        if not permit:
            return JsonResponse({'error': 'Permit not found'}, status=404)
        return JsonResponse({'permit': permit})
    elif request.method == 'PUT':
        try:
            data = json.loads(request.body)
            permit = update_permit(permit_id, data)
            if not permit:
                return JsonResponse({'error': 'Permit not found'}, status=404)
            return JsonResponse({'permit': permit})
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=400)
    elif request.method == 'DELETE':
        if delete_permit(permit_id):
            return JsonResponse({'message': 'Permit deleted'})
        return JsonResponse({'error': 'Permit not found'}, status=404)


@csrf_exempt
@require_http_methods(["GET", "POST"])
@login_required
def hse_api_records(request):
    err = _require_hse_permission(request, 'edit' if request.method == 'POST' else 'view')
    if err:
        return err
    if request.method == 'GET':
        records = get_records()
        return JsonResponse({'records': records})
    elif request.method == 'POST':
        try:
            data = json.loads(request.body)
            record = create_record(data)
            return JsonResponse({'record': record}, status=201)
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=400)


@csrf_exempt
@require_http_methods(["GET", "PUT", "DELETE"])
@login_required
def hse_api_record_detail(request, record_id):
    err = _require_hse_permission(request, 'view' if request.method == 'GET' else 'edit')
    if err:
        return err
    if request.method == 'GET':
        record = get_record(record_id)
        if not record:
            return JsonResponse({'error': 'Record not found'}, status=404)
        return JsonResponse({'record': record})
    elif request.method == 'PUT':
        try:
            data = json.loads(request.body)
            record = update_record(record_id, data)
            if not record:
                return JsonResponse({'error': 'Record not found'}, status=404)
            return JsonResponse({'record': record})
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=400)
    elif request.method == 'DELETE':
        if delete_record(record_id):
            return JsonResponse({'message': 'Record deleted'})
        return JsonResponse({'error': 'Record not found'}, status=404)
