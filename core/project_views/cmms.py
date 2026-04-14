"""
project_views/cmms.py -- Project-scoped CMMS wrappers and hub views.
"""
from __future__ import annotations

from urllib.parse import urlencode

from django.shortcuts import redirect, render

from ..auth_utils import has_permission
from ..cmms_ptw_utils import is_cmms_permit, list_permits
from ..cmms_utils import get_activities
from ..project_data import handover_list
from ..views import get_user, login_required
from .base import _guard, _pctx


def _project_cmms_url(country_id: str, project_id: str, category: str = 'maintenance') -> str:
    return f'/p/{country_id}/{project_id}/{category}/cmms/'


def _project_cmms_permits_url(country_id: str, project_id: str, category: str = 'maintenance') -> str:
    return f'{_project_cmms_url(country_id, project_id, category)}permits/'


def _redirect_with_back(path: str, back_url: str, **params):
    query = {key: value for key, value in params.items() if value not in (None, '')}
    if back_url:
        query['back'] = back_url
    return redirect(f'{path}?{urlencode(query)}' if query else path)


@login_required
def project_cmms_hub(request, country_id, project_id, category='maintenance'):
    country, project = _guard(request, country_id, project_id)
    if not country:
        return redirect('/')

    user = get_user(request)
    if not any(has_permission(user, module_id, 'view') for module_id in ('activities', 'permits', 'handover')):
        return redirect('/')

    cmms_permits = [permit for permit in list_permits() if is_cmms_permit(permit)]
    handovers = handover_list(country_id, project_id)
    active_statuses = {'active', 'pending_closure_issuer', 'pending_closure_hse'}

    ctx, _, _ = _pctx(request, country_id, project_id, {
        'total_activities': len(get_activities()),
        'total_permits': len(cmms_permits),
        'active_permits': sum(1 for permit in cmms_permits if permit.get('status') in active_statuses),
        'total_handovers': len(handovers),
        'open_handovers': sum(1 for handover in handovers if handover.get('status') != 'submitted'),
    }, category=category)
    return render(request, 'core/modules/cmms/hub.html', ctx)


@login_required
def project_cmms_activities(request, country_id, project_id, category='maintenance'):
    country, project = _guard(request, country_id, project_id, 'activities', 'view')
    if not country:
        return redirect('/')
    return _redirect_with_back('/cmms/', _project_cmms_url(country_id, project_id, category))


@login_required
def project_cmms_permits(request, country_id, project_id, category='maintenance'):
    country, project = _guard(request, country_id, project_id, 'permits', 'view')
    if not country:
        return redirect('/')
    return _redirect_with_back(
        '/cmms/ptw/',
        _project_cmms_url(country_id, project_id, category),
        status=request.GET.get('status', '').strip(),
    )


@login_required
def project_cmms_permit_new(request, country_id, project_id, category='maintenance'):
    country, project = _guard(request, country_id, project_id, 'activities', 'view')
    if not country:
        return redirect('/')
    return _redirect_with_back('/cmms/', _project_cmms_url(country_id, project_id, category))


@login_required
def project_cmms_permit_detail(request, country_id, project_id, category='maintenance', permit_id=''):
    country, project = _guard(request, country_id, project_id, 'permits', 'view')
    if not country:
        return redirect('/')
    return _redirect_with_back(
        f'/cmms/ptw/{permit_id}/',
        _project_cmms_permits_url(country_id, project_id, category),
        mode=request.GET.get('mode', '').strip(),
    )
