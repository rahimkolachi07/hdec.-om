"""
project_data/store.py - Store module data for a project.

Storage:
  projects_data/<cid>/<pid>/store/items.json
  media/store/<cid>/<pid>/<item_id>/*
"""
import uuid
from datetime import datetime
from pathlib import Path

from django.conf import settings

from .base import get_module_dir, load_json, save_json


def _path(cid, pid):
    return get_module_dir(cid, pid, 'store') / 'items.json'


def _media_dir(cid, pid, item_id):
    path = Path(settings.MEDIA_ROOT) / 'store' / cid / pid / item_id
    path.mkdir(parents=True, exist_ok=True)
    return path


def _save_pictures(cid, pid, item_id, pictures):
    saved = []
    for picture in pictures or []:
        ext = Path(picture.name or '').suffix.lower() or '.jpg'
        name = f'{uuid.uuid4().hex[:12]}{ext}'
        target = _media_dir(cid, pid, item_id) / name
        with open(target, 'wb') as fh:
            for chunk in picture.chunks():
                fh.write(chunk)
        saved.append(f"{settings.MEDIA_URL}store/{cid}/{pid}/{item_id}/{name}")
    return saved


def store_load(cid, pid) -> list:
    raw = load_json(_path(cid, pid))
    return raw if isinstance(raw, list) else []


def store_get(cid, pid, item_id):
    for item in store_load(cid, pid):
        if item.get('id') == item_id:
            return item
    return None


def store_create(cid, pid, payload, pictures=None):
    items = store_load(cid, pid)
    item_id = str(uuid.uuid4())[:12]
    item = {
        'id': item_id,
        'equipment_name': payload.get('equipment_name', '').strip(),
        'date': payload.get('date', '').strip(),
        'details': payload.get('details', '').strip(),
        'quantity': payload.get('quantity', '').strip(),
        'status': payload.get('status', 'given').strip().lower() or 'given',
        'pictures': _save_pictures(cid, pid, item_id, pictures),
        'created_at': datetime.now().isoformat(),
        'updated_at': datetime.now().isoformat(),
    }
    items.insert(0, item)
    save_json(_path(cid, pid), items)
    return item


def store_update(cid, pid, item_id, payload, pictures=None):
    items = store_load(cid, pid)
    for item in items:
        if item.get('id') != item_id:
            continue
        item['equipment_name'] = payload.get('equipment_name', item.get('equipment_name', '')).strip()
        item['date'] = payload.get('date', item.get('date', '')).strip()
        item['details'] = payload.get('details', item.get('details', '')).strip()
        item['quantity'] = payload.get('quantity', item.get('quantity', '')).strip()
        item['status'] = payload.get('status', item.get('status', 'given')).strip().lower() or 'given'
        retained = payload.get('retain_pictures')
        if retained is not None:
            item['pictures'] = retained if isinstance(retained, list) else [retained]
        if pictures:
            item.setdefault('pictures', []).extend(_save_pictures(cid, pid, item_id, pictures))
        item['updated_at'] = datetime.now().isoformat()
        save_json(_path(cid, pid), items)
        return item
    return None


def store_delete(cid, pid, item_id):
    items = [item for item in store_load(cid, pid) if item.get('id') != item_id]
    save_json(_path(cid, pid), items)
