"""
project_data/cmms/handover.py — Shift handover data for a project.

Storage: projects_data/<cid>/<pid>/cmms/handover/<handover_id>.json
Each handover entry is a separate file. Listing reads all files in the directory.
"""
import io, uuid, openpyxl
from datetime import datetime
from ..base import get_module_dir, load_json, save_json, excel_style_header


def _handover_dir(cid, pid):
    return get_module_dir(cid, pid, 'cmms', 'handover')


def _ho_path(cid, pid, ho_id: str):
    return _handover_dir(cid, pid) / f'{ho_id}.json'


def ho_load(cid, pid) -> list:
    """Return all handover entries sorted by created_at descending."""
    d = _handover_dir(cid, pid)
    entries = []
    for f in d.glob('*.json'):
        data = load_json(f)
        if isinstance(data, dict):
            entries.append(data)
    return sorted(entries, key=lambda h: h.get('created_at', ''), reverse=True)


def ho_get(cid, pid, ho_id: str) -> dict | None:
    return load_json(_ho_path(cid, pid, ho_id))


def ho_save_one(cid, pid, entry: dict):
    """Save a single handover entry to its own file."""
    save_json(_ho_path(cid, pid, entry['id']), entry)


def ho_create(cid, pid, payload: dict) -> dict:
    item = {
        'id': str(uuid.uuid4())[:12],
        'date': payload.get('date', datetime.now().strftime('%Y-%m-%d')),
        'shift': payload.get('shift', 'Day'),
        'timing': payload.get('timing', ''),
        'shift_incharge': payload.get('shift_incharge', ''),
        'technicians': payload.get('technicians', []),
        'major_alarms': payload.get('major_alarms', ''),
        'equipment_breakdown': payload.get('equipment_breakdown', ''),
        'maintenance_activities': payload.get('maintenance_activities', ''),
        'inverter_faults': payload.get('inverter_faults', ''),
        'scb_faults': payload.get('scb_faults', ''),
        'spare_parts': payload.get('spare_parts', ''),
        'key_issues': payload.get('key_issues', ''),
        'pending_work': payload.get('pending_work', ''),
        'instructions_next_shift': payload.get('instructions_next_shift', ''),
        'observation_text': payload.get('observation_text', ''),
        'status': payload.get('status', 'draft'),
        'created_at': datetime.now().isoformat(),
        'created_by': payload.get('created_by', ''),
        'shift_engineer_sig': payload.get('shift_engineer_sig'),
        'incoming_engineer_sig': payload.get('incoming_engineer_sig'),
        'submitted_at': payload.get('submitted_at'),
    }
    if item['status'] == 'submitted' and not item['submitted_at']:
        item['submitted_at'] = datetime.now().isoformat()
    ho_save_one(cid, pid, item)
    return item


def ho_update(cid, pid, ho_id: str, fields: dict) -> dict | None:
    entry = ho_get(cid, pid, ho_id)
    if not entry:
        return None
    entry.update(fields)
    ho_save_one(cid, pid, entry)
    return entry


def ho_delete(cid, pid, ho_id: str):
    p = _ho_path(cid, pid, ho_id)
    if p.exists():
        p.unlink()


# ── Excel Export ───────────────────────────────────────────────────────────────

def ho_export_excel(cid, pid) -> bytes:
    entries = ho_load(cid, pid)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Handover'
    excel_style_header(ws,
        ['Date', 'Shift', 'Shift Incharge', 'Major Alarms', 'Equipment Breakdown',
         'Key Issues', 'Pending Work', 'Status', 'Created At'],
        {'A': 14, 'B': 10, 'C': 22, 'D': 30, 'E': 30, 'F': 30},
    )
    for h in entries:
        ws.append([
            h.get('date', ''), h.get('shift', ''), h.get('shift_incharge', ''),
            h.get('major_alarms', ''), h.get('equipment_breakdown', ''),
            h.get('key_issues', ''), h.get('pending_work', ''),
            h.get('status', ''), (h.get('created_at', '') or '')[:10],
        ])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
