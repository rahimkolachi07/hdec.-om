"""
project_views — Project-scoped module views package.

URL patterns in urls.py import this package as `project_views` and
access views via attribute lookup (e.g. project_views.project_manpower),
so everything is re-exported here.
"""
from .manpower import project_manpower, project_manpower_api
from .store import project_store, project_store_api
from .cmms import (
    project_cmms_hub,
    project_activities, project_activities_api,
    project_permits, project_permit_detail, project_permits_api,
    project_handovers, project_handover_detail, project_handover_api,
)

__all__ = [
    'project_manpower', 'project_manpower_api',
    'project_store', 'project_store_api',
    'project_cmms_hub',
    'project_activities', 'project_activities_api',
    'project_permits', 'project_permit_detail', 'project_permits_api',
    'project_handovers', 'project_handover_detail', 'project_handover_api',
]
