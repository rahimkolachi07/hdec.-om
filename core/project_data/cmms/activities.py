"""
project_data/cmms/activities.py — Activities data for a project.

Storage: projects_data/<cid>/<pid>/cmms/activities.json
"""
import io, uuid, openpyxl
from datetime import datetime
from ..base import get_module_dir, load_json, save_json, excel_style_header


def _path(cid, pid):
    return get_module_dir(cid, pid, 'cmms') / 'activities.json'


def act_load(cid, pid) -> list:
    r = load_json(_path(cid, pid))
    return r if isinstance(r, list) else []


def act_save(cid, pid, data: list):
    save_json(_path(cid, pid), data)


def act_create(cid, pid, payload: dict) -> dict:
    activities = act_load(cid, pid)
    item = {
        'id': str(uuid.uuid4())[:12],
        'name': payload.get('name', ''),
        'equipment': payload.get('equipment', ''),
        'location': payload.get('location', ''),
        'frequency': payload.get('frequency', 'Monthly'),
        'type': payload.get('type', 'PM'),
        'description': payload.get('description', ''),
        'checklist': payload.get('checklist', []),
        'created_at': datetime.now().isoformat(),
        'created_by': payload.get('created_by', ''),
        'status': 'active',
    }
    activities.append(item)
    act_save(cid, pid, activities)
    return item


def act_update(cid, pid, act_id: str, fields: dict):
    activities = act_load(cid, pid)
    for a in activities:
        if a['id'] == act_id:
            for k, v in fields.items():
                if k != 'id':
                    a[k] = v
    act_save(cid, pid, activities)


def act_delete(cid, pid, act_id: str):
    act_save(cid, pid, [a for a in act_load(cid, pid) if a['id'] != act_id])


def act_get(cid, pid, act_id: str) -> dict | None:
    for a in act_load(cid, pid):
        if a['id'] == act_id:
            return a
    return None


# ── Excel Import ───────────────────────────────────────────────────────────────

def act_parse_excel(file_bytes: bytes) -> list:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(h or '').strip().lower() for h in rows[0]]
    col = {h: i for i, h in enumerate(header)}

    def g(row, key, default=''):
        i = col.get(key)
        return str(row[i] or '').strip() if i is not None and i < len(row) else default

    items = []
    for row in rows[1:]:
        if not any(row):
            continue
        items.append({
            'name': g(row, 'name') or g(row, 'activity'),
            'equipment': g(row, 'equipment'),
            'location': g(row, 'location'),
            'type': g(row, 'type') or 'PM',
            'frequency': g(row, 'frequency') or 'Monthly',
            'description': g(row, 'description') or g(row, 'remarks'),
            'checklist': [],
        })
    return [a for a in items if a['name']]


# ── Excel Export ───────────────────────────────────────────────────────────────

def act_export_excel(cid, pid) -> bytes:
    activities = act_load(cid, pid)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Activities'
    excel_style_header(ws,
        ['Name', 'Equipment', 'Location', 'Type', 'Frequency', 'Description', 'Status', 'Created At'],
        {'A': 28, 'B': 22, 'C': 20, 'D': 12, 'E': 14, 'F': 40},
    )
    for a in activities:
        ws.append([
            a.get('name', ''), a.get('equipment', ''), a.get('location', ''),
            a.get('type', ''), a.get('frequency', ''), a.get('description', ''),
            a.get('status', ''), (a.get('created_at', '') or '')[:10],
        ])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
