"""
project_views/base.py — Shared helpers for all project-scoped views.
"""
from ..views import get_user, _ctx, login_required
from ..project_utils import get_country, get_project
from ..auth_utils import can_access_project, has_permission

WORK_TYPES = {
    'general':    'General Work',
    'electrical': 'Electrical Work',
    'mechanical': 'Mechanical Work',
    'hot_work':   'Hot Work',
    'confined':   'Confined Space',
    'height':     'Work at Height',
    'excavation': 'Excavation',
    'lifting':    'Lifting Operations',
}

PERMIT_STATUSES = {
    'pending_issue':     'Pending Issuance',
    'pending_hse':       'Pending HSE',
    'active':            'Active',
    'waiting_for_close': 'Waiting for Close',
    'closed':            'Closed',
    'cancelled':         'Cancelled',
}


def _pctx(request, country_id, project_id, extra=None, category='maintenance'):
    """Build render context with country/project/category loaded."""
    country = get_country(country_id)
    project = get_project(country_id, project_id)
    ctx = _ctx(request, {
        'country': country,
        'project': project,
        'country_id': country_id,
        'project_id': project_id,
        'category': category,
    })
    if extra:
        ctx.update(extra)
    return ctx, country, project


def _guard(request, country_id, project_id, module_id: str = None, level: str = 'view'):
    """Return (country, project) or (None, None) if not found / not allowed."""
    country = get_country(country_id)
    project = get_project(country_id, project_id)
    if not country or not project:
        return None, None
    user = get_user(request)
    if not can_access_project(user, country_id, project_id):
        return None, None
    if module_id and not has_permission(user, module_id, level):
        return None, None
    return country, project
