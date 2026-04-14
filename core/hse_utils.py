"""
HSE data management utilities.

This module stores HSE permits and related records on local disk under
``data/hse`` so the data is shared across browsers and devices.
"""
from __future__ import annotations

import json
import os
import tempfile
import time
import uuid
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

from django.conf import settings

if os.name == 'nt':
    import msvcrt
else:
    import fcntl

HSE_DATA_DIR = Path(
    getattr(settings, 'HSE_DATA_DIR', Path(__file__).resolve().parent.parent / 'data' / 'hse')
).resolve()
PERMITS_FILE = HSE_DATA_DIR / 'permits.json'
RECORDS_FILE = HSE_DATA_DIR / 'records.json'


def _now_iso() -> str:
    return datetime.now().isoformat(timespec='seconds')


def _ensure_storage() -> None:
    HSE_DATA_DIR.mkdir(parents=True, exist_ok=True)
    for file in (PERMITS_FILE, RECORDS_FILE):
        if not file.exists():
            _atomic_write_json(file, [])


def _normalize_payload(data: dict) -> dict:
    if not isinstance(data, dict):
        raise ValueError('Expected a JSON object payload.')
    return dict(data)


def _load(file: Path) -> list:
    _ensure_storage()
    try:
        with file.open('r', encoding='utf-8') as handle:
            data = json.load(handle)
            return data if isinstance(data, list) else []
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return []


def _atomic_write_json(file: Path, data: list) -> None:
    file.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(prefix=f'{file.stem}_', suffix='.tmp', dir=file.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as handle:
            json.dump(data, handle, indent=2, ensure_ascii=False)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_path, file)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def _lock(file: Path):
    file.parent.mkdir(parents=True, exist_ok=True)
    handle = file.open('a+', encoding='utf-8')
    try:
        handle.seek(0, os.SEEK_END)
        if handle.tell() == 0:
            handle.write('0')
            handle.flush()
        handle.seek(0)
        if os.name == 'nt':
            while True:
                try:
                    msvcrt.locking(handle.fileno(), msvcrt.LK_LOCK, 1)
                    break
                except OSError:
                    time.sleep(0.05)
        else:
            fcntl.flock(handle.fileno(), fcntl.LOCK_EX)
        return handle
    except Exception:
        handle.close()
        raise


def _unlock(handle) -> None:
    try:
        handle.seek(0)
        if os.name == 'nt':
            msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
    finally:
        handle.close()


@contextmanager
def _collection_lock(file: Path):
    handle = _lock(file.with_suffix(f'{file.suffix}.lock'))
    try:
        yield
    finally:
        _unlock(handle)


def _get_item(file: Path, item_id: str) -> dict | None:
    for item in _load(file):
        if item.get('id') == item_id:
            return item
    return None


def _create_item(file: Path, data: dict) -> dict:
    payload = _normalize_payload(data)
    with _collection_lock(file):
        items = _load(file)
        now = _now_iso()
        item = {
            **payload,
            'id': str(uuid.uuid4()),
            'created_at': now,
            'updated_at': now,
        }
        items.append(item)
        _atomic_write_json(file, items)
        return item


def _update_item(file: Path, item_id: str, data: dict) -> dict | None:
    payload = _normalize_payload(data)
    with _collection_lock(file):
        items = _load(file)
        now = _now_iso()
        for index, item in enumerate(items):
            if item.get('id') != item_id:
                continue
            updated = {
                **item,
                **payload,
                'id': item_id,
                'created_at': item.get('created_at', now),
                'updated_at': now,
            }
            items[index] = updated
            _atomic_write_json(file, items)
            return updated
    return None


def _delete_item(file: Path, item_id: str) -> bool:
    with _collection_lock(file):
        items = _load(file)
        for index, item in enumerate(items):
            if item.get('id') != item_id:
                continue
            del items[index]
            _atomic_write_json(file, items)
            return True
    return False


def get_permits() -> list:
    return _load(PERMITS_FILE)


def get_permit(permit_id: str) -> dict | None:
    return _get_item(PERMITS_FILE, permit_id)


def create_permit(data: dict) -> dict:
    return _create_item(PERMITS_FILE, data)


def update_permit(permit_id: str, data: dict) -> dict | None:
    return _update_item(PERMITS_FILE, permit_id, data)


def delete_permit(permit_id: str) -> bool:
    return _delete_item(PERMITS_FILE, permit_id)


def get_records() -> list:
    return _load(RECORDS_FILE)


def get_record(record_id: str) -> dict | None:
    return _get_item(RECORDS_FILE, record_id)


def create_record(data: dict) -> dict:
    return _create_item(RECORDS_FILE, data)


def update_record(record_id: str, data: dict) -> dict | None:
    return _update_item(RECORDS_FILE, record_id, data)


def delete_record(record_id: str) -> bool:
    return _delete_item(RECORDS_FILE, record_id)


_ensure_storage()
