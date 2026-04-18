"""
File-backed data layer for administration modules.

Data is stored in admin_data/ at the project root. Records may be used
globally by the control panel or scoped to a specific country/project pair
for the project Administration workspace.
"""
import json
import uuid
from datetime import datetime, timezone
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent / 'admin_data'
BASE.mkdir(exist_ok=True)

VEHICLES_FILE = BASE / 'vehicles.json'
RESIDENCES_FILE = BASE / 'residences.json'
WORKFORCE_FILE = BASE / 'workforce.json'
GATEPASSES_FILE = BASE / 'gate_passes.json'
EQUIPMENT_FILE = BASE / 'equipment.json'
TRAININGS_FILE = BASE / 'trainings.json'


def _now():
    return datetime.now(timezone.utc).isoformat()


def _uid():
    return str(uuid.uuid4())


def _load(path):
    if not path.exists():
        return []
    try:
        with open(path, 'r', encoding='utf-8') as handle:
            return json.load(handle)
    except Exception:
        return []


def _save(path, data):
    with open(path, 'w', encoding='utf-8') as handle:
        json.dump(data, handle, indent=2, ensure_ascii=False)


def _scope_values(country_id=None, project_id=None):
    country = str(country_id or '').strip()
    project = str(project_id or '').strip()
    if not country or not project:
        return '', '', ''
    return country, project, f'{country}/{project}'


def _apply_scope(record, country_id=None, project_id=None):
    country, project, project_key = _scope_values(country_id, project_id)
    if not project_key:
        return record
    record['country_id'] = country
    record['project_id'] = project
    record['project_key'] = project_key
    return record


def _matches_scope(record, country_id=None, project_id=None):
    country, project, project_key = _scope_values(country_id, project_id)
    if not project_key:
        return True
    return (
        str(record.get('project_key') or '').strip() == project_key or
        (
            str(record.get('country_id') or '').strip() == country and
            str(record.get('project_id') or '').strip() == project
        )
    )


def _prepare_records(path, country_id=None, project_id=None):
    records = _load(path)
    country, project, project_key = _scope_values(country_id, project_id)
    if not project_key or not records:
        return records

    # Legacy datasets were saved without scope. If the whole file is still
    # unscoped, claim it for the first project that opens it so existing data
    # remains visible after the Administration split.
    if any(str(item.get('project_key') or '').strip() for item in records):
        return records

    for item in records:
        _apply_scope(item, country, project)
    _save(path, records)
    return records


def list_records(path, country_id=None, project_id=None):
    records = _prepare_records(path, country_id, project_id)
    if not _scope_values(country_id, project_id)[2]:
        return records
    return [record for record in records if _matches_scope(record, country_id, project_id)]


def get_record(path, rid, country_id=None, project_id=None):
    records = _prepare_records(path, country_id, project_id)
    return next(
        (
            record for record in records
            if record.get('id') == rid and _matches_scope(record, country_id, project_id)
        ),
        None,
    )


def create_record(path, data, country_id=None, project_id=None):
    records = _prepare_records(path, country_id, project_id)
    record = {**data, 'id': _uid(), 'created_at': _now(), 'updated_at': _now()}
    _apply_scope(record, country_id, project_id)
    records.append(record)
    _save(path, records)
    return record


def update_record(path, rid, data, country_id=None, project_id=None):
    records = _prepare_records(path, country_id, project_id)
    for index, record in enumerate(records):
        if record.get('id') != rid or not _matches_scope(record, country_id, project_id):
            continue
        records[index] = {**record, **data, 'id': rid, 'updated_at': _now()}
        _apply_scope(
            records[index],
            record.get('country_id') or country_id,
            record.get('project_id') or project_id,
        )
        _save(path, records)
        return records[index]
    return None


def delete_record(path, rid, country_id=None, project_id=None):
    records = _prepare_records(path, country_id, project_id)
    new_records = [
        record for record in records
        if not (record.get('id') == rid and _matches_scope(record, country_id, project_id))
    ]
    if len(new_records) < len(records):
        _save(path, new_records)
        return True
    return False


def list_vehicles(country_id=None, project_id=None):
    return list_records(VEHICLES_FILE, country_id, project_id)


def get_vehicle(rid, country_id=None, project_id=None):
    return get_record(VEHICLES_FILE, rid, country_id, project_id)


def create_vehicle(data, country_id=None, project_id=None):
    return create_record(VEHICLES_FILE, data, country_id, project_id)


def update_vehicle(rid, data, country_id=None, project_id=None):
    return update_record(VEHICLES_FILE, rid, data, country_id, project_id)


def delete_vehicle(rid, country_id=None, project_id=None):
    return delete_record(VEHICLES_FILE, rid, country_id, project_id)


def list_residences(country_id=None, project_id=None):
    return list_records(RESIDENCES_FILE, country_id, project_id)


def get_residence(rid, country_id=None, project_id=None):
    return get_record(RESIDENCES_FILE, rid, country_id, project_id)


def create_residence(data, country_id=None, project_id=None):
    return create_record(RESIDENCES_FILE, data, country_id, project_id)


def update_residence(rid, data, country_id=None, project_id=None):
    return update_record(RESIDENCES_FILE, rid, data, country_id, project_id)


def delete_residence(rid, country_id=None, project_id=None):
    return delete_record(RESIDENCES_FILE, rid, country_id, project_id)


def list_workforce(country_id=None, project_id=None):
    return list_records(WORKFORCE_FILE, country_id, project_id)


def get_workforce(rid, country_id=None, project_id=None):
    return get_record(WORKFORCE_FILE, rid, country_id, project_id)


def create_workforce(data, country_id=None, project_id=None):
    return create_record(WORKFORCE_FILE, data, country_id, project_id)


def update_workforce(rid, data, country_id=None, project_id=None):
    return update_record(WORKFORCE_FILE, rid, data, country_id, project_id)


def delete_workforce(rid, country_id=None, project_id=None):
    return delete_record(WORKFORCE_FILE, rid, country_id, project_id)


def list_gatepasses(country_id=None, project_id=None):
    return list_records(GATEPASSES_FILE, country_id, project_id)


def get_gatepass(rid, country_id=None, project_id=None):
    return get_record(GATEPASSES_FILE, rid, country_id, project_id)


def create_gatepass(data, country_id=None, project_id=None):
    return create_record(GATEPASSES_FILE, data, country_id, project_id)


def update_gatepass(rid, data, country_id=None, project_id=None):
    return update_record(GATEPASSES_FILE, rid, data, country_id, project_id)


def delete_gatepass(rid, country_id=None, project_id=None):
    return delete_record(GATEPASSES_FILE, rid, country_id, project_id)


def list_equipment(country_id=None, project_id=None):
    return list_records(EQUIPMENT_FILE, country_id, project_id)


def get_equipment(rid, country_id=None, project_id=None):
    return get_record(EQUIPMENT_FILE, rid, country_id, project_id)


def create_equipment(data, country_id=None, project_id=None):
    return create_record(EQUIPMENT_FILE, data, country_id, project_id)


def update_equipment(rid, data, country_id=None, project_id=None):
    return update_record(EQUIPMENT_FILE, rid, data, country_id, project_id)


def delete_equipment(rid, country_id=None, project_id=None):
    return delete_record(EQUIPMENT_FILE, rid, country_id, project_id)


def list_trainings(country_id=None, project_id=None):
    return list_records(TRAININGS_FILE, country_id, project_id)


def get_training(rid, country_id=None, project_id=None):
    return get_record(TRAININGS_FILE, rid, country_id, project_id)


def create_training(data, country_id=None, project_id=None):
    return create_record(TRAININGS_FILE, data, country_id, project_id)


def update_training(rid, data, country_id=None, project_id=None):
    return update_record(TRAININGS_FILE, rid, data, country_id, project_id)


def delete_training(rid, country_id=None, project_id=None):
    return delete_record(TRAININGS_FILE, rid, country_id, project_id)
