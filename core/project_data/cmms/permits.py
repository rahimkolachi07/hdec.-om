"""
project_data/cmms/permits.py — Permits data for a project.

Storage: projects_data/<cid>/<pid>/cmms/permits/<permit_id>.json
Each permit is a separate file. Listing reads all files in the directory.
"""
import io, uuid, openpyxl
from datetime import datetime
from ..base import get_module_dir, load_json, save_json, excel_style_header


def _permits_dir(cid, pid):
    return get_module_dir(cid, pid, 'cmms', 'permits')


def _permit_path(cid, pid, permit_id: str):
    return _permits_dir(cid, pid) / f'{permit_id}.json'


def permit_load(cid, pid) -> list:
    """Return all permits sorted by created_at descending."""
    d = _permits_dir(cid, pid)
    permits = []
    for f in d.glob('*.json'):
        data = load_json(f)
        if isinstance(data, dict):
            permits.append(data)
    return sorted(permits, key=lambda p: p.get('created_at', ''), reverse=True)


def permit_get(cid, pid, permit_id: str) -> dict | None:
    return load_json(_permit_path(cid, pid, permit_id))


def permit_save_one(cid, pid, permit: dict):
    """Save a single permit to its own file."""
    save_json(_permit_path(cid, pid, permit['id']), permit)


def permit_create(cid, pid, payload: dict) -> dict:
    item = {
        'id': str(uuid.uuid4())[:12],
        'permit_number': None,
        'status': 'pending_issue',
        'receiver': payload.get('receiver', ''),
        'receiver_name': payload.get('receiver_name', ''),
        'issuer': None, 'issuer_name': None,
        'hse_officer': None, 'hse_name': None,
        'job_description': payload.get('job_description', ''),
        'location': payload.get('location', ''),
        'equipment': payload.get('equipment', ''),
        'work_type': payload.get('work_type', 'general'),
        'hazards': payload.get('hazards', []),
        'precautions': payload.get('precautions', []),
        'isolation_required': payload.get('isolation_required', False),
        'isolation_details': payload.get('isolation_details', ''),
        'valid_from': payload.get('valid_from', ''),
        'valid_until': payload.get('valid_until', ''),
        'workers': payload.get('workers', []),
        'receiver_signature': None,
        'issuer_signature': None,
        'hse_signature': None,
        'closure_receiver_signature': None,
        'closure_issuer_signature': None,
        'closure_hse_signature': None,
        'activity_images': [],
        'created_at': datetime.now().isoformat(),
        'issued_at': None,
        'hse_signed_at': None,
        'closed_at': None,
        'closed_by': None,
        'closed_by_name': None,
        'application_datetime': payload.get('application_datetime', datetime.now().isoformat()),
        'comments': '',
    }
    permit_save_one(cid, pid, item)
    return item


def permit_update(cid, pid, permit_id: str, fields: dict) -> dict | None:
    permit = permit_get(cid, pid, permit_id)
    if not permit:
        return None
    permit.update(fields)
    permit_save_one(cid, pid, permit)
    return permit


def permit_delete(cid, pid, permit_id: str):
    p = _permit_path(cid, pid, permit_id)
    if p.exists():
        p.unlink()


# ── Excel Export ───────────────────────────────────────────────────────────────

def permit_export_excel(cid, pid) -> bytes:
    permits = permit_load(cid, pid)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Permits'
    excel_style_header(ws,
        ['Permit No.', 'Equipment', 'Location', 'Work Type', 'Receiver', 'Status', 'Valid From', 'Valid Until', 'Created At'],
        {'A': 16, 'B': 22, 'C': 20, 'D': 14, 'E': 22, 'F': 18},
    )
    for p in permits:
        ws.append([
            p.get('permit_number', 'TBD'), p.get('equipment', ''), p.get('location', ''),
            p.get('work_type', ''), p.get('receiver_name', ''), p.get('status', ''),
            (p.get('valid_from', '') or '')[:16], (p.get('valid_until', '') or '')[:16],
            (p.get('created_at', '') or '')[:10],
        ])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
