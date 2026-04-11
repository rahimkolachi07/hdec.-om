"""
project_data — Per-project data storage package.

Disk layout:
  projects_data/
    <country_id>/
      <project_id>/
        manpower/
          data.json
        cmms/
          activities.json
          permits/
            <permit_id>.json      ← one file per permit
          handover/
            <handover_id>.json    ← one file per handover entry
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
from .cmms import (
    act_load, act_save, act_create, act_update, act_delete, act_get,
    act_parse_excel, act_export_excel,
    permit_load, permit_get, permit_create, permit_update, permit_delete,
    permit_save_one, permit_export_excel,
    ho_load, ho_get, ho_create, ho_update, ho_delete,
    ho_save_one, ho_export_excel,
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
    # activities
    'act_load', 'act_save', 'act_create', 'act_update', 'act_delete', 'act_get',
    'act_parse_excel', 'act_export_excel',
    # permits
    'permit_load', 'permit_get', 'permit_create', 'permit_update', 'permit_delete',
    'permit_save_one', 'permit_export_excel',
    # handover
    'ho_load', 'ho_get', 'ho_create', 'ho_update', 'ho_delete',
    'ho_save_one', 'ho_export_excel',
]
