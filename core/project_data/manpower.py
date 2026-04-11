"""
project_data/manpower.py — Manpower data for a project.

Storage: projects_data/<cid>/<pid>/manpower/data.json
"""
import io, uuid, openpyxl
from datetime import datetime
from .base import get_module_dir, load_json, save_json, excel_style_header


def _path(cid, pid):
    return get_module_dir(cid, pid, 'manpower') / 'data.json'


def mp_load(cid, pid) -> dict:
    raw = load_json(_path(cid, pid))
    if isinstance(raw, dict):
        return raw
    return {'engineers': [], 'technicians': [], 'updated_at': None}


def mp_save(cid, pid, data: dict):
    data['updated_at'] = datetime.now().isoformat()
    save_json(_path(cid, pid), data)


def mp_get(cid, pid) -> dict:
    return mp_load(cid, pid)


def mp_add_person(cid, pid, category: str, name: str, role: str, dept: str = ''):
    data = mp_load(cid, pid)
    data.setdefault(category, []).append({
        'id': str(uuid.uuid4())[:8],
        'name': name.strip(),
        'role': role.strip(),
        'dept': dept.strip(),
        'schedule': {},
    })
    mp_save(cid, pid, data)


def mp_remove_person(cid, pid, category: str, person_id: str):
    data = mp_load(cid, pid)
    data[category] = [p for p in data.get(category, []) if p.get('id') != person_id]
    mp_save(cid, pid, data)


def mp_update_schedule(cid, pid, category: str, person_id: str, schedule: dict):
    data = mp_load(cid, pid)
    for p in data.get(category, []):
        if p.get('id') == person_id:
            p['schedule'] = schedule
            break
    mp_save(cid, pid, data)


def mp_bulk_update(cid, pid, engineers: list, technicians: list):
    mp_save(cid, pid, {'engineers': engineers, 'technicians': technicians})


# ── Excel Import ───────────────────────────────────────────────────────────────

SHIFT_VALUES = {'Day', 'Night', 'General', 'OFF', 'Leave', 'Rest'}


def mp_parse_excel(file_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

    def _parse_sheet(ws) -> list:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        header = rows[0]
        meta_cols, date_cols = [], []
        for i, h in enumerate(header):
            if h is None:
                continue
            h_str = str(h).strip()
            if h_str.lower() in ('name', 'role', 'department', 'dept', 'team'):
                meta_cols.append((i, h_str.lower()))
            else:
                if isinstance(h, datetime):
                    date_cols.append((i, h.strftime('%Y-%m-%d')))
                else:
                    try:
                        from datetime import date as _date
                        if isinstance(h, _date):
                            date_cols.append((i, h.strftime('%Y-%m-%d')))
                        else:
                            from dateutil.parser import parse as _dparse
                            dt = _dparse(str(h_str))
                            date_cols.append((i, dt.strftime('%Y-%m-%d')))
                    except Exception:
                        pass

        people = []
        for row in rows[1:]:
            if not any(row):
                continue
            name = role = dept = ''
            for ci, cname in meta_cols:
                val = str(row[ci] or '').strip()
                if cname == 'name':
                    name = val
                elif cname == 'role':
                    role = val
                elif cname in ('department', 'dept', 'team'):
                    dept = val
            if not name:
                continue
            schedule = {}
            for ci, dstr in date_cols:
                raw = str(row[ci] or '').strip()
                norm = raw.capitalize()
                if norm not in SHIFT_VALUES:
                    norm = 'General'
                schedule[dstr] = norm
            people.append({'id': str(uuid.uuid4())[:8], 'name': name, 'role': role, 'dept': dept, 'schedule': schedule})
        return people

    sheet_names_lower = {s.lower(): s for s in wb.sheetnames}
    engineers = []
    technicians = []

    if 'engineers' in sheet_names_lower:
        engineers = _parse_sheet(wb[sheet_names_lower['engineers']])
    if 'technicians' in sheet_names_lower:
        technicians = _parse_sheet(wb[sheet_names_lower['technicians']])

    if not engineers and not technicians and wb.sheetnames:
        all_people = _parse_sheet(wb[wb.sheetnames[0]])
        for p in all_people:
            if 'tech' in p.get('role', '').lower():
                technicians.append(p)
            else:
                engineers.append(p)

    return {'engineers': engineers, 'technicians': technicians}


# ── Excel Export ───────────────────────────────────────────────────────────────

def mp_export_excel(cid, pid) -> bytes:
    data = mp_load(cid, pid)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for category, label in [('engineers', 'Engineers'), ('technicians', 'Technicians')]:
        people = data.get(category, [])
        ws = wb.create_sheet(label)
        all_dates = sorted({d for p in people for d in p.get('schedule', {})})
        excel_style_header(ws, ['Name', 'Role', 'Department'] + all_dates,
                           {'A': 24, 'B': 22, 'C': 20})
        for p in people:
            row = [p.get('name', ''), p.get('role', ''), p.get('dept', '')]
            row += [p.get('schedule', {}).get(d, '') for d in all_dates]
            ws.append(row)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
