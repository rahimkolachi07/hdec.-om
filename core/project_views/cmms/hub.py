"""
project_views/cmms/hub.py — CMMS hub view for a project.
"""
from django.shortcuts import render, redirect
from ..base import _pctx, _guard
from ...views import login_required
from ...project_data import act_load, permit_load, ho_load


@login_required
def project_cmms_hub(request, country_id, project_id, category='maintenance'):
    country, project = _guard(request, country_id, project_id)
    if not country:
        return redirect('/')

    acts = act_load(country_id, project_id)
    permits = permit_load(country_id, project_id)
    handovers = ho_load(country_id, project_id)

    active_permits = [p for p in permits if p.get('status') in ('active', 'waiting_for_close')]
    open_handovers = [h for h in handovers if h.get('status') == 'draft']

    ctx, _, _ = _pctx(request, country_id, project_id, {
        'total_activities': len(acts),
        'total_permits': len(permits),
        'active_permits': len(active_permits),
        'total_handovers': len(handovers),
        'open_handovers': len(open_handovers),
    }, category=category)
    return render(request, 'core/modules/cmms/hub.html', ctx)
