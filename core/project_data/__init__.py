"""
project_data — Per-project data storage package.

Disk layout:
  projects_data/
    <country_id>/
      <project_id>/
        manpower/
          data.json
        store/
          data.json
"""
from .base import get_project_dir, get_module_dir, load_json, save_json, get_blank_template
from .manpower import (
    mp_load, mp_save, mp_get,
    mp_add_person, mp_remove_person, mp_update_schedule, mp_bulk_update,
    mp_parse_excel, mp_export_excel,
)
from .store import (
    store_load, store_get, store_create, store_update, store_delete,
)
from .handover import (
    handover_list, handover_get, handover_find_by_date_shift,
    handover_create, handover_update, handover_delete, handover_export_excel,
)

__all__ = [
    # base
    'get_project_dir', 'get_module_dir', 'get_blank_template',
    # manpower
    'mp_load', 'mp_save', 'mp_get',
    'mp_add_person', 'mp_remove_person', 'mp_update_schedule', 'mp_bulk_update',
    'mp_parse_excel', 'mp_export_excel',
    # store
    'store_load', 'store_get', 'store_create', 'store_update', 'store_delete',
    # handover
    'handover_list', 'handover_get', 'handover_find_by_date_shift',
    'handover_create', 'handover_update', 'handover_delete', 'handover_export_excel',
]
