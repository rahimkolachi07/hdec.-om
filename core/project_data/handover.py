"""
project_data/handover.py -- Shift handover data for a project.

Storage:
  projects_data/<country_id>/<project_id>/cmms/handover/<handover_id>.json
"""
from __future__ import annotations

import io
import uuid
from datetime import datetime

import openpyxl

from .base import excel_style_header, get_module_dir, load_json, save_json

VALID_SHIFTS = {'Day', 'Night', 'General'}
VALID_STATUSES = {'draft', 'submitted'}
TEXT_FIELDS = (
    'date',
    'shift',
    'timing',
    'shift_incharge',
    'major_alarms',
    'equipment_breakdown',
    'maintenance_activities',
    'inverter_faults',
    'scb_faults',
    'spare_parts',
    'key_issues',
    'pending_work',
    'instructions_next_shift',
    'observation_text',
    'shift_engineer_sig',
    'incoming_engineer_sig',
    'submitted_at',
)


def _dir(cid: str, pid: str):
    return get_module_dir(cid, pid, 'cmms', 'handover')


def _path(cid: str, pid: str, handover_id: str):
    return _dir(cid, pid) / f'{handover_id}.json'


def _clean_text(value) -> str:
    return str(value or '').strip()


def _clean_list(values) -> list[str]:
    cleaned = []
    seen = set()
    for value in values or []:
        item = _clean_text(value)
        if not item:
            continue
        key = item.lower()
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(item)
    return cleaned


def _normalize_status(value: str) -> str:
    return 'submitted' if str(value or '').strip().lower() == 'submitted' else 'draft'


def _normalize_shift(value: str) -> str:
    shift = _clean_text(value) or 'General'
    return shift if shift in VALID_SHIFTS else 'General'


def _sort_key(handover: dict) -> tuple:
    shift_order = {'Night': 2, 'Day': 1, 'General': 0}
    return (
        handover.get('date', ''),
        shift_order.get(handover.get('shift', 'General'), 0),
        handover.get('updated_at') or handover.get('created_at') or '',
        handover.get('id', ''),
    )


def _normalize_handover(raw: dict | None, handover_id: str = '') -> dict | None:
    if not isinstance(raw, dict):
        return None
    handover = {
        'id': _clean_text(raw.get('id') or handover_id or str(uuid.uuid4())),
        'date': _clean_text(raw.get('date')),
        'shift': _normalize_shift(raw.get('shift')),
        'timing': _clean_text(raw.get('timing')),
        'shift_incharge': _clean_text(raw.get('shift_incharge')),
        'technicians': _clean_list(raw.get('technicians', [])),
        'status': _normalize_status(raw.get('status')),
        'created_at': _clean_text(raw.get('created_at')),
        'created_by': _clean_text(raw.get('created_by')),
        'updated_at': _clean_text(raw.get('updated_at')),
        'updated_by': _clean_text(raw.get('updated_by')),
        'submitted_at': _clean_text(raw.get('submitted_at')),
        'deleted_at': _clean_text(raw.get('deleted_at')),
        'deleted_by': _clean_text(raw.get('deleted_by')),
    }
    for field in TEXT_FIELDS:
        if field in handover:
            continue
        handover[field] = _clean_text(raw.get(field))
    if handover['status'] != 'submitted':
        handover['submitted_at'] = ''
    return handover


def handover_list(cid: str, pid: str) -> list[dict]:
    items = []
    for path in sorted(_dir(cid, pid).glob('*.json')):
        handover = _normalize_handover(load_json(path), path.stem)
        if handover and not handover.get('deleted_at'):
            items.append(handover)
    return sorted(items, key=_sort_key, reverse=True)


def handover_get(cid: str, pid: str, handover_id: str, include_deleted: bool = False) -> dict | None:
    handover = _normalize_handover(load_json(_path(cid, pid, handover_id)), handover_id)
    if handover and (include_deleted or not handover.get('deleted_at')):
        return handover
    return None


def handover_find_by_date_shift(cid: str, pid: str, date_value: str, shift: str) -> dict | None:
    clean_date = _clean_text(date_value)
    clean_shift = _normalize_shift(shift)
    for handover in handover_list(cid, pid):
        if handover.get('date') == clean_date and handover.get('shift') == clean_shift:
            return handover
    return None


def handover_create(cid: str, pid: str, data: dict, created_by: str = '') -> dict:
    handover_id = str(uuid.uuid4())
    now = datetime.now().isoformat()
    handover = _normalize_handover({
        **(data or {}),
        'id': handover_id,
        'created_at': now,
        'created_by': _clean_text(created_by),
        'updated_at': now,
        'updated_by': _clean_text(created_by),
    }, handover_id) or {}
    if handover.get('status') == 'submitted' and not handover.get('submitted_at'):
        handover['submitted_at'] = now
    save_json(_path(cid, pid, handover_id), handover)
    return handover


def handover_update(cid: str, pid: str, handover_id: str, updates: dict, updated_by: str = '') -> dict | None:
    current = handover_get(cid, pid, handover_id)
    if not current:
        return None

    payload = dict(current)
    for field in TEXT_FIELDS:
        if field in updates:
            payload[field] = _clean_text(updates.get(field))
    if 'date' in updates:
        payload['date'] = _clean_text(updates.get('date'))
    if 'shift' in updates:
        payload['shift'] = _normalize_shift(updates.get('shift'))
    if 'timing' in updates:
        payload['timing'] = _clean_text(updates.get('timing'))
    if 'shift_incharge' in updates:
        payload['shift_incharge'] = _clean_text(updates.get('shift_incharge'))
    if 'technicians' in updates:
        payload['technicians'] = _clean_list(updates.get('technicians', []))
    if 'status' in updates:
        payload['status'] = _normalize_status(updates.get('status'))

    payload['updated_at'] = datetime.now().isoformat()
    payload['updated_by'] = _clean_text(updated_by)
    if payload.get('status') == 'submitted' and not payload.get('submitted_at'):
        payload['submitted_at'] = payload['updated_at']
    if payload.get('status') != 'submitted':
        payload['submitted_at'] = ''

    handover = _normalize_handover(payload, handover_id)
    save_json(_path(cid, pid, handover_id), handover)
    return handover


def handover_delete(cid: str, pid: str, handover_id: str, deleted_by: str = '') -> bool:
    current = handover_get(cid, pid, handover_id, include_deleted=True)
    if not current:
        return False
    current['deleted_at'] = datetime.now().isoformat()
    current['deleted_by'] = _clean_text(deleted_by)
    current['updated_at'] = current['deleted_at']
    current['updated_by'] = current['deleted_by']
    save_json(_path(cid, pid, handover_id), current)
    return True


def handover_export_excel(cid: str, pid: str) -> bytes:
    handovers = handover_list(cid, pid)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Shift Handover'
    excel_style_header(
        ws,
        [
            'Date',
            'Shift',
            'Timing',
            'Shift In-Charge',
            'Technicians',
            'Status',
            'Major Alarms',
            'Equipment Breakdown',
            'Maintenance Activities',
            'Inverter Faults',
            'SCB Faults',
            'Spare Parts',
            'Key Issues',
            'Pending Work',
            'Instructions Next Shift',
            'Observation Text',
            'Created At',
            'Created By',
            'Submitted At',
        ],
        {
            'A': 14, 'B': 12, 'C': 18, 'D': 22, 'E': 26, 'F': 12,
            'G': 24, 'H': 24, 'I': 30, 'J': 22, 'K': 22, 'L': 20,
            'M': 26, 'N': 24, 'O': 30, 'P': 30, 'Q': 20, 'R': 18, 'S': 20,
        },
    )
    for handover in handovers:
        ws.append([
            handover.get('date', ''),
            handover.get('shift', ''),
            handover.get('timing', ''),
            handover.get('shift_incharge', ''),
            ', '.join(handover.get('technicians', [])),
            handover.get('status', '').capitalize(),
            handover.get('major_alarms', ''),
            handover.get('equipment_breakdown', ''),
            handover.get('maintenance_activities', ''),
            handover.get('inverter_faults', ''),
            handover.get('scb_faults', ''),
            handover.get('spare_parts', ''),
            handover.get('key_issues', ''),
            handover.get('pending_work', ''),
            handover.get('instructions_next_shift', ''),
            handover.get('observation_text', ''),
            handover.get('created_at', ''),
            handover.get('created_by', ''),
            handover.get('submitted_at', ''),
        ])

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
