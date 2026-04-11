from .hub import project_cmms_hub
from .activities import project_activities, project_activities_api
from .permits import project_permits, project_permit_detail, project_permits_api
from .handover import project_handovers, project_handover_detail, project_handover_api

__all__ = [
    'project_cmms_hub',
    'project_activities', 'project_activities_api',
    'project_permits', 'project_permit_detail', 'project_permits_api',
    'project_handovers', 'project_handover_detail', 'project_handover_api',
]
