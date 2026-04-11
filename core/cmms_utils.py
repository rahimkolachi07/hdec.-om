"""
CMMS Data Management Utilities
Handles activities, checklist records, and permits using JSON file storage.
"""
import json, uuid, os, zipfile, io, math
from pathlib import Path
from datetime import datetime
from django.conf import settings

# ── Data directories ──────────────────────────────────────────────────────
CMMS_DATA_DIR = getattr(settings, 'CMMS_DATA_DIR', Path(__file__).resolve().parent.parent / 'cmms_data')
CMMS_DATA_DIR.mkdir(exist_ok=True)

MEDIA_ROOT = getattr(settings, 'MEDIA_ROOT', Path(__file__).resolve().parent.parent / 'media')
PHOTOS_DIR = MEDIA_ROOT / 'cmms' / 'photos'
CHECKLISTS_DIR = MEDIA_ROOT / 'cmms' / 'checklists'
PERMIT_TEMPLATES_DIR = MEDIA_ROOT / 'cmms' / 'permit_templates'
HANDOVER_IMAGES_DIR = MEDIA_ROOT / 'cmms' / 'handover_images'

for d in [PHOTOS_DIR, CHECKLISTS_DIR, PERMIT_TEMPLATES_DIR, HANDOVER_IMAGES_DIR]:
    d.mkdir(parents=True, exist_ok=True)

ACTIVITIES_FILE = CMMS_DATA_DIR / 'activities.json'
RECORDS_FILE    = CMMS_DATA_DIR / 'records.json'
PERMITS_FILE    = CMMS_DATA_DIR / 'permits.json'
HANDOVERS_FILE  = CMMS_DATA_DIR / 'handovers.json'


# ── Generic JSON helpers ──────────────────────────────────────────────────
def _load(file: Path) -> list:
    if not file.exists():
        return []
    try:
        with open(file, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []


def _save(file: Path, data: list):
    with open(file, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, default=str)


# ── ACTIVITIES ────────────────────────────────────────────────────────────
def get_activities() -> list:
    return _load(ACTIVITIES_FILE)


def get_activity(activity_id: str) -> dict | None:
    return next((a for a in get_activities() if a['id'] == activity_id), None)


def get_activities_for_user(username: str, role: str) -> list:
    """
    Return activities visible to this user.
    - admin: all activities
    - technician: activities where assigned_technician matches OR unassigned (anyone)
    - engineer roles: activities where assigned_engineer matches OR unassigned (anyone)
    - Unassigned means the activity is open to whoever is on duty.
    """
    activities = get_activities()
    if role == 'admin':
        return activities
    if role == 'technician':
        return [a for a in activities
                if not a.get('assigned_technician') or a.get('assigned_technician') == username]
    # maintenance_engineer, operation_engineer, hse_engineer, viewer
    return [a for a in activities
            if not a.get('assigned_engineer') or a.get('assigned_engineer') == username]


def get_activities_for_month(month: str) -> list:
    """Return activities whose month OR scheduled_date month matches."""
    return [a for a in get_activities()
            if a.get('month', '').startswith(month)
            or a.get('scheduled_date', '').startswith(month)]


def create_activity(data: dict) -> dict:
    activities = get_activities()
    scheduled_date = data.get('scheduled_date', '')
    month = data.get('month', '') or (scheduled_date[:7] if scheduled_date else datetime.now().strftime('%Y-%m'))
    # Parse pm_count (Schedule PM — how many units/blocks/times per period)
    pm_raw = data.get('pm_count', data.get('scheduled_pm', ''))
    try:
        pm_count = int(str(pm_raw).strip()) if str(pm_raw).strip().isdigit() else 1
    except Exception:
        pm_count = 1

    activity = {
        'id':                   str(uuid.uuid4()),
        'month':                month,
        'scheduled_date':       scheduled_date,
        'name':                 data.get('name', ''),
        'equipment':            data.get('equipment', ''),
        'location':             data.get('location', ''),
        'frequency':            data.get('frequency', 'once'),
        'pm_count':             pm_count,   # units/blocks per period
        'assigned_engineer':    data.get('assigned_engineer', ''),
        'assigned_technician':  data.get('assigned_technician', ''),
        'checklist_items':      data.get('checklist_items', []),
        'checklist_file':       data.get('checklist_file', ''),
        'notes':                data.get('notes', ''),
        'created_at':           datetime.now().isoformat(),
        'created_by':           data.get('created_by', 'admin'),
    }
    activities.append(activity)
    _save(ACTIVITIES_FILE, activities)
    return activity


def update_activity(activity_id: str, updates: dict) -> dict | None:
    activities = get_activities()
    for i, a in enumerate(activities):
        if a['id'] == activity_id:
            activities[i].update(updates)
            _save(ACTIVITIES_FILE, activities)
            return activities[i]
    return None


def delete_activity(activity_id: str):
    activities = [a for a in get_activities() if a['id'] != activity_id]
    _save(ACTIVITIES_FILE, activities)


# ── CHECKLIST RECORDS ─────────────────────────────────────────────────────
def get_records() -> list:
    return _load(RECORDS_FILE)


def get_record(record_id: str) -> dict | None:
    return next((r for r in get_records() if r['id'] == record_id), None)


def get_records_for_activity(activity_id: str) -> list:
    return [r for r in get_records() if r['activity_id'] == activity_id]


def get_or_create_record(activity_id: str, date: str, engineer: str, engineer_name: str) -> dict:
    """Get existing record for an activity+date, or create a new one."""
    records = get_records()
    for r in records:
        if r['activity_id'] == activity_id and r['date'] == date and r['engineer'] == engineer:
            return r
    activity = get_activity(activity_id)
    record = {
        'id':               str(uuid.uuid4()),
        'activity_id':      activity_id,
        'activity_name':    activity['name'] if activity else '',
        'date':             date,
        'engineer':         engineer,
        'engineer_name':    engineer_name,
        'technician':       activity.get('assigned_technician', '') if activity else '',
        'checkpoints':      {},   # {item_id: True/False}
        'remarks':          {},   # {item_id: "remark text"}
        'engineer_signature': None,
        'before_photos':    [],
        'after_photos':     [],
        'completed':        False,
        'completed_at':     None,
        'created_at':       datetime.now().isoformat(),
    }
    records.append(record)
    _save(RECORDS_FILE, records)
    return record


def _activity_due_today(act: dict, date_str: str) -> bool:
    """
    Return True if the activity should generate a record on date_str based on its frequency.
    """
    freq = act.get('frequency', 'once').lower()
    scheduled_date = act.get('scheduled_date', '')
    month = act.get('month', '')

    if freq == 'once':
        # Only on the exact scheduled date
        return scheduled_date == date_str

    if freq == 'daily':
        # Every day in the activity's month
        return date_str.startswith(month) if month else True

    if freq == 'weekly':
        # Every Monday (or start of week) within the activity's month
        from datetime import date as _date
        try:
            d = _date.fromisoformat(date_str)
            return d.weekday() == 0 and date_str.startswith(month)  # Mondays
        except Exception:
            return False

    if freq == 'monthly':
        # First day of the month
        try:
            return date_str.endswith('-01') and date_str.startswith(month[:4])
        except Exception:
            return False

    if freq in ('quarterly', 'half_yearly', 'annual'):
        # Just once; treat scheduled_date or first day of month
        if scheduled_date:
            return scheduled_date == date_str
        return date_str == (month + '-01') if month else False

    # fallback: match scheduled_date
    return scheduled_date == date_str if scheduled_date else date_str.startswith(month)


def auto_create_daily_records(date_str: str = None) -> list:
    """
    For all activities that are due today (based on frequency), auto-create
    records for duty engineers. Returns list of records touched.
    Called on hub/activities page load so engineers always see today's work.
    """
    if not date_str:
        date_str = datetime.now().strftime('%Y-%m-%d')
    month = date_str[:7]  # 'YYYY-MM'

    activities = get_activities_for_month(month)
    if not activities:
        return []

    duty = get_duty_staff(date_str)
    duty_engineers = duty.get('engineers', [])

    # Build name→username map from user accounts for matching schedule names
    from .auth_utils import get_all_users
    all_users = get_all_users()
    name_to_user = {u['name'].strip().lower(): u for u in all_users}

    touched = []
    for act in activities:
        if not _activity_due_today(act, date_str):
            continue

        assigned_eng = act.get('assigned_engineer', '').strip()

        if assigned_eng:
            user = name_to_user.get(assigned_eng.lower()) or next(
                (u for u in all_users if u['username'] == assigned_eng), None)
            eng_username = user['username'] if user else assigned_eng
            eng_name = user['name'] if user else assigned_eng
            rec = get_or_create_record(act['id'], date_str, eng_username, eng_name)
            touched.append(rec)
        else:
            if not duty_engineers:
                rec = get_or_create_record(act['id'], date_str, 'duty_engineer', 'Duty Engineer')
                touched.append(rec)
            else:
                for eng in duty_engineers:
                    eng_name = eng.get('name', '')
                    matched = name_to_user.get(eng_name.strip().lower())
                    eng_username = matched['username'] if matched else eng_name
                    rec = get_or_create_record(act['id'], date_str, eng_username, eng_name)
                    touched.append(rec)

    return touched


def import_activities_from_schedule(month: str = None) -> dict:
    """
    Read media/schedule.xlsx and create CMMS activities for the given month.
    Deduplicates by task name — skips tasks already imported for this month.
    Returns {'created': [...], 'skipped': int}
    """
    if not month:
        month = datetime.now().strftime('%Y-%m')

    xl_path = MEDIA_ROOT / 'schedule.xlsx'
    if not xl_path.exists():
        raise FileNotFoundError(f'Schedule file not found: {xl_path}')

    import openpyxl
    wb = openpyxl.load_workbook(str(xl_path), data_only=True)
    ws = wb.active

    # Parse unique task entries
    import re as _re
    seen_names = set()
    tasks = []

    # Keywords that indicate header/summary rows — not real tasks
    SKIP_KEYWORDS = {
        'equipment', 'total', 'compliance score', 'pending', 'maintenance description',
        'task description', 'description', 'week 1', 'week 2', 'week 3', 'week 4', 'week 5',
        'area', 'block number', 'performing date', 'checklist', 'performed', 'status',
        'none', '', 'nan',
    }

    # Detect frequency from task name suffix:  "NAME - D", "NAME-D", "NAME - W", etc.
    def _detect_freq(name_upper: str) -> str:
        # Match patterns like "- D", "-D", "- W", "-W", "- M", "-M", "- HY", "-HY" at end
        m = _re.search(r'[-\s]+([A-Z]{1,2})\s*$', name_upper.strip())
        if m:
            suffix = m.group(1)
            return {'D': 'daily', 'W': 'weekly', 'M': 'monthly', 'HY': 'half_yearly'}.get(suffix, 'once')
        return 'once'

    for row in ws.iter_rows(values_only=True):
        area_raw = row[0]
        name_raw = row[1]
        scheduled_pm = row[2]

        area = str(area_raw).strip() if area_raw is not None else ''
        name = str(name_raw).strip() if name_raw is not None else ''

        if not name or name.lower() in SKIP_KEYWORDS:
            continue
        # Skip header-like rows (area is "AREA" or task looks like a date range)
        if area.lower() == 'area':
            continue
        if _re.search(r'week\s+\d', name.lower()):
            continue

        name_upper = name.upper()
        freq = _detect_freq(name_upper)

        # Deduplicate by normalised task name
        norm = name_upper.strip()
        if norm in seen_names:
            continue
        seen_names.add(norm)

        # Clean up area
        clean_area = area if area.lower() not in ('', 'none', 'nan') else ''

        tasks.append({
            'name': name.title(),
            'area': clean_area,
            'frequency': freq,
            'scheduled_pm': scheduled_pm,
        })

    # Load existing activities to skip duplicates for this month
    existing = get_activities_for_month(month)
    existing_names_lower = {a['name'].strip().lower() for a in existing}

    created = []
    skipped = 0
    for task in tasks:
        if task['name'].strip().lower() in existing_names_lower:
            skipped += 1
            continue

        # For daily tasks start from 1st of month; others from 1st too
        act = create_activity({
            'month': month,
            'scheduled_date': '',   # frequency-based, not single date
            'name': task['name'],
            'equipment': task['name'],
            'location': task['area'],
            'frequency': task['frequency'],
            'assigned_engineer': '',
            'assigned_technician': '',
            'checklist_items': [],
            'pm_count': task['scheduled_pm'],
            'notes': f"Imported from schedule.xlsx — PM scheduled: {task['scheduled_pm']}",
            'created_by': 'system_import',
        })
        created.append(act)

    return {'created': created, 'skipped': skipped}


def get_today_tasks_for_dashboard(date_str: str = None) -> list:
    """
    Smart scheduling engine — returns what tasks must be done today.

    Rules per frequency:
      daily      → every single day, 1 task
      weekly     → fixed weekday per task type; SVG 6 units spread Mon–Sat
      monthly    → distribute pm_count blocks across month days (1 block/day or more)
      half_yearly→ ARCS Robots: 20/day cycling; others: 2+ blocks/day cycling through month
      once       → only on scheduled_date
    """
    from datetime import date as _date
    import calendar as _cal

    if not date_str:
        date_str = datetime.now().strftime('%Y-%m-%d')

    d = _date.fromisoformat(date_str)
    weekday      = d.weekday()           # 0=Mon … 6=Sun
    dom          = d.day                 # 1–31
    month        = date_str[:7]
    days_in_month = _cal.monthrange(d.year, d.month)[1]
    day_of_year  = d.timetuple().tm_yday

    activities = get_activities_for_month(month)

    # Records already created today → status per activity
    all_records = get_records()
    completed_ids   = {r['activity_id'] for r in all_records if r.get('date') == date_str and r.get('completed')}
    in_progress_ids = {r['activity_id'] for r in all_records if r.get('date') == date_str}

    # ── Fixed weekday assignments for weekly tasks ─────────────────────────
    # weekday 0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri, 5=Sat, 6=Sun
    WEEKLY_FIXED_DAY = {
        'edg test run - ss':              0,  # Monday
        'edg test run - sb':              1,  # Tuesday
        'weather monitoring station - w': 2,  # Wednesday
    }

    def _pm(act):
        """Get integer pm_count from activity."""
        raw = act.get('pm_count', act.get('scheduled_pm', 1))
        try:
            return max(1, int(str(raw).strip())) if str(raw).strip().isdigit() else 1
        except Exception:
            return 1

    def _block_range(pm, dom, days_in_month):
        """
        Distribute pm blocks across month days.
        Returns (block_start, block_end) for today, or None if today has no block.
        """
        if pm <= 0:
            return None
        bpd = max(1, math.ceil(pm / days_in_month))  # blocks per day
        start = (dom - 1) * bpd + 1
        end   = min(start + bpd - 1, pm)
        if start > pm:
            return None
        return (start, end)

    def _status(act_id):
        if act_id in completed_ids:   return 'done'
        if act_id in in_progress_ids: return 'in_progress'
        return 'pending'

    def _freq_label(f):
        return {'daily': 'D', 'weekly': 'W', 'monthly': 'M', 'half_yearly': 'HY', 'once': '1x'}.get(f, f)

    tasks = []

    for act in activities:
        freq   = act.get('frequency', 'once').lower()
        name_l = act['name'].lower().strip()
        pm     = _pm(act)
        block_info = ''
        due = False

        # ── DAILY ─────────────────────────────────────────────────────────
        if freq == 'daily':
            due = True
            block_info = 'All areas'

        # ── WEEKLY ────────────────────────────────────────────────────────
        elif freq == 'weekly':
            if 'svg cooling' in name_l:
                # 6 SVG units → Mon(1) Tue(2) Wed(3) Thu(4) Fri(5) Sat(6), Sun off
                if weekday <= 5:
                    unit = weekday + 1
                    due = True
                    block_info = f'SVG Unit {unit}'
            else:
                # Find fixed day
                assigned = next((day for pat, day in WEEKLY_FIXED_DAY.items()
                                 if pat in name_l), 0)  # default Monday
                if weekday == assigned:
                    due = True

        # ── MONTHLY ───────────────────────────────────────────────────────
        elif freq == 'monthly':
            if pm > 1:
                # Multi-block task — distribute blocks across days
                rng = _block_range(pm, dom, days_in_month)
                if rng:
                    due = True
                    s, e = rng
                    if 'hvac' in name_l:
                        block_info = f'Unit {s}' if s == e else f'Units {s}–{e}'
                    elif 'power station' in name_l or 'substation panel' in name_l:
                        block_info = f'MVPS Block {s}' if s == e else f'MVPS Block {s}–{e}'
                    else:
                        block_info = f'Block {s}' if s == e else f'Block {s}–{e}'
            else:
                # Single occurrence — hash-assign to a stable day this month
                assigned_dom = (abs(hash(name_l)) % days_in_month) + 1
                if dom == assigned_dom:
                    due = True

        # ── HALF-YEARLY ───────────────────────────────────────────────────
        elif freq == 'half_yearly':
            if 'arcs robots' in name_l or 'robot box' in name_l:
                # 20 robots per day, cycling through total continuously
                robots_per_day = 20
                cycle_days = max(1, math.ceil(pm / robots_per_day))
                day_in_cycle = (day_of_year - 1) % cycle_days
                start = day_in_cycle * robots_per_day + 1
                end   = min(start + robots_per_day - 1, pm)
                due = True
                block_info = f'Robots {start}–{end} / {pm}'
            else:
                # SCB / MVPS HY — distribute pm blocks across month (2+ per day)
                rng = _block_range(pm, dom, days_in_month)
                if rng:
                    due = True
                    s, e = rng
                    if 'scb' in name_l:
                        block_info = f'SCB {s}' if s == e else f'SCB {s}–{e}'
                    else:
                        block_info = f'Block {s}' if s == e else f'Block {s}–{e}'

        # ── ONCE ──────────────────────────────────────────────────────────
        elif freq == 'once':
            sd = act.get('scheduled_date', '')
            if sd == date_str:
                due = True

        if due:
            tasks.append({
                'activity_id':       act['id'],
                'name':              act['name'],
                'location':          act.get('location', ''),
                'frequency':         freq,
                'freq_label':        _freq_label(freq),
                'block_info':        block_info,
                'pm_count':          pm,
                'checklist_required': len(act.get('checklist_items', [])) > 0,
                'status':            _status(act['id']),
            })

    # Sort: daily → weekly → monthly → HY → once
    _order = {'daily': 0, 'weekly': 1, 'monthly': 2, 'half_yearly': 3, 'once': 4}
    tasks.sort(key=lambda t: _order.get(t['frequency'], 5))
    return tasks


def get_today_records_for_user(username: str, role: str, date_str: str = None) -> list:
    """Return all records for today that belong to this user."""
    if not date_str:
        date_str = datetime.now().strftime('%Y-%m-%d')
    records = get_records()
    today_records = [r for r in records if r.get('date') == date_str]
    if role == 'admin':
        return today_records
    if role == 'technician':
        return [r for r in today_records if r.get('technician') == username]
    return [r for r in today_records if r.get('engineer') == username]


def get_all_technicians_from_schedule(date_str: str = None) -> list:
    """
    Read schedule_store.json and return ALL technicians with their shift for the given date.
    Returns list of dicts: {name, shift, on_duty: bool}
    Duty technicians come first, sorted by name.
    """
    from django.conf import settings as _s
    schedule_file = Path(_s.BASE_DIR) / 'schedule_store.json'
    if not date_str:
        date_str = datetime.now().strftime('%Y-%m-%d')
    try:
        with open(schedule_file, 'r', encoding='utf-8') as f:
            store = json.load(f)
    except Exception:
        return []

    OFF_SHIFTS = {'r', 'rest', 'off', 'leave', 'eid leave', 'on leave', 'annual leave', ''}

    result = []
    for t in store.get('technicians', []):
        name = t.get('name', '').strip()
        if not name:
            continue
        shift = t.get('schedule', {}).get(date_str, '').strip()
        on_duty = shift.lower() not in OFF_SHIFTS
        result.append({
            'name':     name,
            'shift':    shift or 'OFF',
            'on_duty':  on_duty,
        })

    # Sort: on-duty first, then alphabetical
    result.sort(key=lambda x: (not x['on_duty'], x['name'].lower()))
    return result


def update_record(record_id: str, updates: dict) -> dict | None:
    records = get_records()
    for i, r in enumerate(records):
        if r['id'] == record_id:
            records[i].update(updates)
            _save(RECORDS_FILE, records)
            return records[i]
    return None


# ── PHOTO MANAGEMENT ──────────────────────────────────────────────────────
def save_photo(record_id: str, phase: str, uploaded_file) -> str:
    """
    Save uploaded photo file.
    phase: 'before' or 'after'
    Returns relative path from MEDIA_ROOT.
    """
    folder = PHOTOS_DIR / record_id / phase
    folder.mkdir(parents=True, exist_ok=True)
    ext = Path(uploaded_file.name).suffix.lower() or '.jpg'
    filename = f"photo_{datetime.now().strftime('%H%M%S_%f')}{ext}"
    file_path = folder / filename
    with open(file_path, 'wb') as f:
        for chunk in uploaded_file.chunks():
            f.write(chunk)
    return f"cmms/photos/{record_id}/{phase}/{filename}"


def delete_photo(record_id: str, phase: str, filename: str):
    file_path = PHOTOS_DIR / record_id / phase / filename
    if file_path.exists():
        file_path.unlink()


# ── ZIP GENERATION ────────────────────────────────────────────────────────
def generate_record_zip(record_id: str) -> io.BytesIO | None:
    """
    Create a ZIP with:
      before/  — all before photos
      after/   — all after photos
      checklist.json — checklist data
    """
    record = get_record(record_id)
    if not record:
        return None

    buffer = io.BytesIO()
    activity = get_activity(record['activity_id'])
    folder_name = f"{record.get('activity_name', 'activity')}_{record['date']}".replace(' ', '_')

    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Add before photos
        before_dir = PHOTOS_DIR / record_id / 'before'
        if before_dir.exists():
            for photo in sorted(before_dir.iterdir()):
                zf.write(photo, f"{folder_name}/before/{photo.name}")

        # Add after photos
        after_dir = PHOTOS_DIR / record_id / 'after'
        if after_dir.exists():
            for photo in sorted(after_dir.iterdir()):
                zf.write(photo, f"{folder_name}/after/{photo.name}")

        # Add checklist JSON
        checklist_data = {
            'activity': activity['name'] if activity else record.get('activity_name', ''),
            'equipment': activity['equipment'] if activity else '',
            'location': activity.get('location', '') if activity else '',
            'date': record['date'],
            'engineer': record['engineer_name'],
            'completed': record['completed'],
            'completed_at': record.get('completed_at'),
            'items': [],
        }
        if activity:
            for item in activity.get('checklist_items', []):
                item_id = str(item['id'])
                checklist_data['items'].append({
                    'section': item.get('section', ''),
                    'description': item.get('description', ''),
                    'checked': record['checkpoints'].get(item_id, False),
                    'remark': record['remarks'].get(item_id, ''),
                })
        zf.writestr(f"{folder_name}/checklist.json", json.dumps(checklist_data, indent=2))

        # Also add the uploaded checklist PDF if present
        if activity and activity.get('checklist_file'):
            pdf_path = MEDIA_ROOT / activity['checklist_file']
            if pdf_path.exists():
                zf.write(pdf_path, f"{folder_name}/checklist_template{pdf_path.suffix}")

    buffer.seek(0)
    return buffer


def generate_activity_month_zip(activity_id: str) -> io.BytesIO | None:
    """
    Combined ZIP for all records of an activity:
      before/  — all before photos from all records
      after/   — all after photos from all records
      checklists/  — per-date checklist JSONs
    """
    activity = get_activity(activity_id)
    if not activity:
        return None
    records = get_records_for_activity(activity_id)

    buffer = io.BytesIO()
    folder_name = f"{activity['name']}_{activity['month']}".replace(' ', '_')

    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for record in records:
            date_prefix = record['date']

            before_dir = PHOTOS_DIR / record['id'] / 'before'
            if before_dir.exists():
                for photo in sorted(before_dir.iterdir()):
                    zf.write(photo, f"{folder_name}/before/{date_prefix}_{photo.name}")

            after_dir = PHOTOS_DIR / record['id'] / 'after'
            if after_dir.exists():
                for photo in sorted(after_dir.iterdir()):
                    zf.write(photo, f"{folder_name}/after/{date_prefix}_{photo.name}")

            checklist_data = {
                'date': record['date'],
                'engineer': record['engineer_name'],
                'completed': record['completed'],
                'checkpoints': record['checkpoints'],
                'remarks': record['remarks'],
            }
            zf.writestr(
                f"{folder_name}/checklists/{date_prefix}_checklist.json",
                json.dumps(checklist_data, indent=2)
            )

    buffer.seek(0)
    return buffer


# ── PERMITS ───────────────────────────────────────────────────────────────
PERMIT_STATUSES = {
    'pending_issue':    'Pending Issuance',
    'pending_hse':      'Pending HSE Sign-off',
    'active':           'Active / Approved',
    'waiting_for_close': 'Waiting for Close',
    'closed':           'Closed',
    'cancelled':        'Cancelled',
}

WORK_TYPES = {
    'cold_work':       'Cold Work',
    'hot_work':        'Hot Work',
    'electrical':      'Electrical Work',
    'confined_space':  'Confined Space Entry',
    'mechanical':      'Mechanical Work',
    'civil':           'Civil Work',
    'at_height':       'Work at Height',
}


def get_permits() -> list:
    return _load(PERMITS_FILE)


def get_permit(permit_id: str) -> dict | None:
    return next((p for p in get_permits() if p['id'] == permit_id), None)


def get_permits_for_user(username: str, role: str) -> list:
    permits = get_permits()
    if role in ('admin', 'operation_engineer', 'hse_engineer'):
        return permits
    return [p for p in permits if p.get('receiver') == username]


def create_permit(data: dict) -> dict:
    permits = get_permits()
    permit = {
        'id':                    str(uuid.uuid4()),
        'permit_number':         None,
        'isolation_cert_number': None,
        'status':                'pending_issue',
        # People
        'receiver':              data.get('receiver', ''),
        'receiver_name':         data.get('receiver_name', ''),
        'receiver_company':      data.get('receiver_company', 'POWERCHINA'),
        'receiver_id':           data.get('receiver_id', ''),
        'issuer':                None,
        'issuer_name':           None,
        'hse_officer':           None,
        'hse_name':              None,
        # Work info
        'job_description':       data.get('job_description', ''),
        'location':              data.get('location', ''),
        'equipment':             data.get('equipment', ''),
        'work_type':             data.get('work_type', 'electrical'),
        'tools_equipment':       data.get('tools_equipment', ''),
        'expected_duration':     data.get('expected_duration', ''),
        'num_employees':         data.get('num_employees', ''),
        'sld_drawing_no':        data.get('sld_drawing_no', ''),
        # Energy state
        'energized_lines':       data.get('energized_lines', False),
        'de_energized_lines':    data.get('de_energized_lines', True),
        # Risk checkboxes
        'risks': data.get('risks', {
            'electrocution': False, 'arc_flash': False, 'flying_particles': False,
            'noise': False, 'falling_objects': False, 'protruding_objects': False,
            'tripping_slipping': False, 'electric_shock': False, 'fire': False,
            'manual_handling': False, 'electric_burn': False, 'near_overhead_lines': False,
            'other_risk': '',
        }),
        # Documents to attach
        'docs_to_attach': data.get('docs_to_attach', {
            'method_statement': False, 'risk_assessment': False, 'other_doc': '',
        }),
        # Precaution checklist (Yes/No/NA)
        'precaution_checks': data.get('precaution_checks', {
            'safe_distance':       'N/A',
            'safe_distance_voltage': '',
            'safe_distance_dist':  '',
            'loto_required':       'No',
            'confined_space':      'No',
            'power_isolated':      'Yes',
            'isolation_type_switch': False,
            'isolation_type_loto':   False,
            'num_locks':           '',
            'lines_de_energized':  'Yes',
            'tools_tested':        'Yes',
            'other_precaution':    '',
        }),
        # Inspected areas
        'inspected_areas': data.get('inspected_areas', {
            'fire_ext_type': '', 'fire_ext_qty': '', 'fire_ext_size': '',
            'access_escape': False, 'danger_sign': False, 'lighting': False,
            'safety_barriers': False, 'stick': False, 'portable_radio': False,
            'other_area': '',
        }),
        # PPE
        'ppe_required': data.get('ppe_required', {
            'helmet': False, 'safety_shoes': False, 'elec_gloves': False,
            'elec_gloves_rating': '', 'half_mask': False, 'safety_goggles': False,
            'reflective_vest': False, 'dust_mask': False, 'safety_clothes': False,
            'face_shield': False, 'arc_flash_ppe': False, 'ear_plugs': False,
            'other_ppe': '',
        }),
        # Hazards & precaution text
        'hazards':               data.get('hazards', ''),
        'precautions':           data.get('precautions', ''),
        'additional_precautions': data.get('additional_precautions', ['', '', '', '', '']),
        # Isolation
        'isolation_required':    data.get('isolation_required', False),
        'isolation_details':     data.get('isolation_details', ''),
        'isolation_type': data.get('isolation_type', {
            'electrical': False, 'mechanical': False, 'earthing': False, 'others': '',
        }),
        'isolation_sequence':    data.get('isolation_sequence', []),
        'de_isolation_sequence': data.get('de_isolation_sequence', []),
        # Validity
        'valid_from':            data.get('valid_from', ''),
        'valid_until':           data.get('valid_until', ''),
        # Workers list (name, iqama for sign-in table)
        'workers_list':          data.get('workers_list', []),
        'workers':               data.get('workers', ''),
        # Closure
        'closure': data.get('closure', {
            'work_completed': False, 'work_incomplete': False,
            'tools_removed': False, 'loto_closed': False,
            'permit_suspended': False, 'housekeeping_done': False,
        }),
        # Signatures — Work Started section
        'receiver_signature':    data.get('receiver_signature'),
        'issuer_signature':      None,
        'hse_signature':         None,
        # Signatures — Closure section (collected when closing permit)
        'closure_receiver_signature': None,
        'closure_issuer_signature':   None,
        'closure_hse_signature':      None,
        # Who closed the permit
        'closed_by':             None,
        'closed_by_name':        None,
        # Activity images (uploaded at close step)
        'activity_images':       [],
        # Application date/time (editable by submitter)
        'application_datetime':  data.get('application_datetime', datetime.now().isoformat()),
        # Timestamps
        'created_at':            datetime.now().isoformat(),
        'issued_at':             None,
        'hse_signed_at':         None,
        'closed_at':             None,
        'comments':              [],
        # Email sent flags
        'email_sent_to_issuer':  False,
        'email_sent_to_hse':     False,
        'email_sent_to_engineers': False,
    }
    permits.append(permit)
    _save(PERMITS_FILE, permits)
    return permit


def update_permit(permit_id: str, updates: dict) -> dict | None:
    permits = get_permits()
    for i, p in enumerate(permits):
        if p['id'] == permit_id:
            permits[i].update(updates)
            _save(PERMITS_FILE, permits)
            return permits[i]
    return None


def get_next_permit_number() -> str:
    permits = get_permits()
    year = datetime.now().year
    existing = [
        p['permit_number'] for p in permits
        if p.get('permit_number') and f'PTW-{year}' in p['permit_number']
    ]
    seq = len(existing) + 1
    return f"PTW-{year}-{seq:04d}"


def get_next_isolation_number() -> str:
    permits = get_permits()
    year = datetime.now().year
    existing = [
        p['isolation_cert_number'] for p in permits
        if p.get('isolation_cert_number') and f'ISO-{year}' in p['isolation_cert_number']
    ]
    seq = len(existing) + 1
    return f"ISO-{year}-{seq:04d}"


# ── MANPOWER HELPERS ─────────────────────────────────────────────────────
def get_duty_staff(date_str: str = None) -> dict:
    """
    Read schedule_store.json and return engineers + technicians on duty
    for a given date (default: today). Excludes REST/Leave/OFF.
    Returns {'engineers': [...], 'technicians': [...]}
    """
    from django.conf import settings as _s
    schedule_file = Path(_s.BASE_DIR) / 'schedule_store.json'
    if not date_str:
        date_str = datetime.now().strftime('%Y-%m-%d')
    try:
        with open(schedule_file, 'r', encoding='utf-8') as f:
            store = json.load(f)
    except Exception:
        return {'engineers': [], 'technicians': []}

    OFF_SHIFTS = {'r', 'rest', 'off', 'leave', 'eid leave', 'on leave', 'annual leave', ''}

    engineers = []
    for e in store.get('engineers', []):
        shift = e.get('schedule', {}).get(date_str, '')
        if shift.strip().lower() not in OFF_SHIFTS:
            engineers.append({
                'name': e.get('name', ''),
                'role': e.get('role', ''),
                'shift': shift,
            })

    technicians = []
    for t in store.get('technicians', []):
        shift = t.get('schedule', {}).get(date_str, '')
        if shift.strip().lower() not in OFF_SHIFTS:
            technicians.append({
                'name': t.get('name', ''),
                'shift': shift,
            })

    return {'engineers': engineers, 'technicians': technicians}


# ── PERMIT PDF (MP-10 FORM 3 — EXACT FORMAT) ──────────────────────────────
def generate_permit_pdf(permit_id: str) -> io.BytesIO | None:
    """Generate PDF matching the exact MP-10 FORM 3 Electrical Work Permit format."""
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        from reportlab.lib.units import mm
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Table, TableStyle,
            Spacer, PageBreak, Image as RLImage, KeepTogether
        )
        from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
        import base64, tempfile
    except ImportError:
        return None

    permit = get_permit(permit_id)
    if not permit:
        return None

    buffer = io.BytesIO()
    LM = RM = 12*mm
    TM = BM = 12*mm
    W = A4[0] - LM - RM

    doc = SimpleDocTemplate(buffer, pagesize=A4,
        leftMargin=LM, rightMargin=RM, topMargin=TM, bottomMargin=BM)

    # ── Shared helpers ──────────────────────────────────────────────────────
    BLK   = colors.black
    RED   = colors.HexColor('#c00000')
    GRAY  = colors.HexColor('#f2f2f2')
    LGRAY = colors.HexColor('#d9d9d9')

    def P(txt, size=8, bold=False, align=TA_LEFT, color=BLK):
        fs = size
        fn = 'Helvetica-Bold' if bold else 'Helvetica'
        return Paragraph(str(txt or ''), ParagraphStyle('_p',
            fontName=fn, fontSize=fs, textColor=color,
            alignment=align, leading=fs*1.3, spaceBefore=1, spaceAfter=1))

    def CB(checked):
        return '■' if checked else '□'

    def sig_img(b64_str, w=45*mm, h=14*mm):
        if not b64_str:
            return P('')
        try:
            data = b64_str.split(',', 1)[1] if ',' in b64_str else b64_str
            raw = base64.b64decode(data)
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
            tmp.write(raw); tmp.close()
            img = RLImage(tmp.name, width=w, height=h)
            os.unlink(tmp.name)
            return img
        except Exception:
            return P('(signed)')

    TS = TableStyle  # shortcut

    def border_ts(extra=None):
        s = [
            ('BOX',       (0,0), (-1,-1), 0.5, BLK),
            ('INNERGRID', (0,0), (-1,-1), 0.5, BLK),
            ('TOPPADDING',    (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ('LEFTPADDING',   (0,0), (-1,-1), 3),
            ('RIGHTPADDING',  (0,0), (-1,-1), 3),
            ('VALIGN',    (0,0), (-1,-1), 'MIDDLE'),
        ]
        if extra:
            s.extend(extra)
        return TS(s)

    def hdr_ts():
        return border_ts([('BACKGROUND', (0,0), (-1,0), GRAY)])

    # shorthand values
    p   = permit
    ptw = p.get('permit_number') or ''
    iso = p.get('isolation_cert_number') or ''
    risks   = p.get('risks', {})
    prec    = p.get('precaution_checks', {})
    insp    = p.get('inspected_areas', {})
    ppe_d   = p.get('ppe_required', {})
    docs_d  = p.get('docs_to_attach', {})
    workers_list = p.get('workers_list', [])
    closure = p.get('closure', {})
    add_prec= p.get('additional_precautions', ['','','','',''])
    valid_from = str(p.get('valid_from',''))
    start_date = valid_from[:10] if valid_from else ''
    start_time = valid_from[11:16] if len(valid_from) > 10 else ''

    elements = []

    # ═══════════════════════════════════════════════════════════════════════
    # PAGE 1
    # ═══════════════════════════════════════════════════════════════════════

    # ── HEADER TABLE ────────────────────────────────────────────────────────
    hdr_data = [[
        P('POWERCHINA', 10, bold=True),
        P('ELECTRICAL-PERMIT TO WOK', 14, bold=True, align=TA_CENTER),
        P('SANA TAIBAH', 10, bold=True, align=TA_RIGHT),
    ]]
    hdr = Table(hdr_data, colWidths=[W*0.2, W*0.6, W*0.2])
    hdr.setStyle(border_ts([
        ('SPAN', (0,0), (0,0)),
        ('ALIGN', (0,0), (0,0), 'LEFT'),
    ]))
    elements.append(hdr)

    sub_hdr_data = [[
        P('MP-10 FORM 3 Electrical Work Permit', 8, align=TA_CENTER),
    ],[
        P('1100MW Al Henakiyah Solar Photovoltaic Independent Power Plant', 9, bold=True, align=TA_CENTER),
    ]]
    sub_hdr = Table(sub_hdr_data, colWidths=[W])
    sub_hdr.setStyle(border_ts())
    elements.append(sub_hdr)
    elements.append(Spacer(1, 1*mm))

    # ── ROW 1: PTW Ref, Starting Date, Time ─────────────────────────────────
    r1 = Table([[
        P('PTW Ref. No:', 8, bold=True), P(ptw, 9),
        P('Starting Date', 8, bold=True), P(start_date, 9),
        P('Time', 8, bold=True), P(start_time, 9),
    ]], colWidths=[W*0.12, W*0.22, W*0.12, W*0.22, W*0.08, W*0.24])
    r1.setStyle(border_ts([('BACKGROUND',(0,0),(0,0),GRAY),('BACKGROUND',(2,0),(2,0),GRAY),('BACKGROUND',(4,0),(4,0),GRAY)]))
    elements.append(r1)

    # ── ROW 2: Company, Expected Duration, Time ──────────────────────────────
    r2 = Table([[
        P('Company\nName:', 8, bold=True), P(p.get('receiver_company','POWERCHINA'), 9),
        P('Expected\nDuration', 8, bold=True), P(p.get('expected_duration',''), 9),
        P('Time', 8, bold=True), P('', 9),
    ]], colWidths=[W*0.12, W*0.22, W*0.12, W*0.22, W*0.08, W*0.24])
    r2.setStyle(border_ts([('BACKGROUND',(0,0),(0,0),GRAY),('BACKGROUND',(2,0),(2,0),GRAY),('BACKGROUND',(4,0),(4,0),GRAY)]))
    elements.append(r2)

    # ── ROW 3: Isolation Cert, Number of Employees ───────────────────────────
    r3 = Table([[
        P('Isolation\nCertificate #', 8, bold=True), P(iso, 9),
        P('Number of Employees', 8, bold=True), P(p.get('num_employees',''), 9),
    ]], colWidths=[W*0.12, W*0.22, W*0.34, W*0.32])
    r3.setStyle(border_ts([('BACKGROUND',(0,0),(0,0),GRAY),('BACKGROUND',(2,0),(2,0),GRAY)]))
    elements.append(r3)

    # ── ROW 4: Energized / De-Energized ──────────────────────────────────────
    r4 = Table([[
        P(f'Energized Lines/Equipment  {CB(p.get("energized_lines",False))}', 8, bold=True),
        P(f'De-Energized Line/ Equipment  {CB(p.get("de_energized_lines",True))}', 8, bold=True),
    ]], colWidths=[W*0.5, W*0.5])
    r4.setStyle(border_ts())
    elements.append(r4)

    # ── Work Description ──────────────────────────────────────────────────────
    wd = Table([
        [P('Work Description:', 8, bold=True)],
        [P(p.get('job_description',''), 9)],
    ], colWidths=[W])
    wd.setStyle(border_ts([('BACKGROUND',(0,0),(-1,0),GRAY), ('MINROWHEIGHT',(0,1),(0,1),14*mm)]))
    elements.append(wd)

    # ── Location ──────────────────────────────────────────────────────────────
    loc = Table([
        [P('Location of job to be performed:', 8, bold=True)],
        [P(p.get('location',''), 9)],
    ], colWidths=[W])
    loc.setStyle(border_ts([('BACKGROUND',(0,0),(-1,0),GRAY), ('MINROWHEIGHT',(0,1),(0,1),10*mm)]))
    elements.append(loc)

    # ── Tools/Equipment ───────────────────────────────────────────────────────
    tools = Table([
        [P("Tool/Equipment's to be used:", 8, bold=True)],
        [P(p.get('tools_equipment',''), 9)],
    ], colWidths=[W])
    tools.setStyle(border_ts([('BACKGROUND',(0,0),(-1,0),GRAY), ('MINROWHEIGHT',(0,1),(0,1),10*mm)]))
    elements.append(tools)

    # ── RISKS ─────────────────────────────────────────────────────────────────
    risk_hdr = Table([[P('Identify risk associated with this Electrical work', 8, bold=True)]],
        colWidths=[W])
    risk_hdr.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),GRAY)]))
    elements.append(risk_hdr)

    def _r(key): return CB(risks.get(key, False))
    risk_rows = [
        [P(f'{_r("electrocution")} Electrocution',8), P(f'{_r("arc_flash")} Arc Flash',8),
         P(f'{_r("flying_particles")} Flying particles',8), P(f'{_r("noise")} Noise',8)],
        [P(f'{_r("falling_objects")} Falling Objects',8), P(f'{_r("protruding_objects")} Protruding objects, parts',8),
         P(f'{_r("tripping_slipping")} Tripping / Slipping',8), P(f'{_r("electric_shock")} Electric shock',8)],
        [P(f'{_r("fire")} Fire',8), P(f'{_r("manual_handling")} Manual handling',8),
         P(f'{_r("electric_burn")} Electric Burn',8), P(f'{_r("near_overhead_lines")} Near Overhead lines',8)],
        [P(f'Other (Specify): {risks.get("other_risk","")}',8), '', '', ''],
    ]
    risk_t = Table(risk_rows, colWidths=[W/4]*4)
    risk_t.setStyle(border_ts([('SPAN',(0,3),(-1,3))]))
    elements.append(risk_t)

    # ── DOCUMENTS TO ATTACH ───────────────────────────────────────────────────
    dt = docs_d
    doc_row = Table([[
        P('The following document shall be attached with this permit', 8, bold=True),
    ]], colWidths=[W])
    doc_row.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),GRAY)]))
    elements.append(doc_row)

    doc_chk = Table([[
        P(f'{CB(dt.get("method_statement",False))} Method Statement', 8),
        P(f'{CB(dt.get("risk_assessment",False))} Risk Assessment', 8),
        P(f'Other (specify): {dt.get("other_doc","")}', 8),
    ]], colWidths=[W*0.28, W*0.28, W*0.44])
    doc_chk.setStyle(border_ts())
    elements.append(doc_chk)

    # ── PRECAUTIONS TABLE ─────────────────────────────────────────────────────
    prec_hdr_row = Table([[
        P('Precautions require you to complete the work safely (Filled by PTW Receiver and verified by PTW issuer)', 8, bold=True),
        P('Yes', 8, bold=True, align=TA_CENTER),
        P('No', 8, bold=True, align=TA_CENTER),
        P('N/A', 8, bold=True, align=TA_CENTER),
    ]], colWidths=[W*0.7, W*0.1, W*0.1, W*0.1])
    prec_hdr_row.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),GRAY)]))
    elements.append(prec_hdr_row)

    def prec_row(label, key, extra=''):
        v = prec.get(key, 'N/A')
        y = CB(v == 'Yes'); n = CB(v == 'No'); na = CB(v == 'N/A')
        return Table([[
            P(label + (' ' + extra if extra else ''), 8),
            P(y, 9, align=TA_CENTER), P(n, 9, align=TA_CENTER), P(na, 9, align=TA_CENTER),
        ]], colWidths=[W*0.7, W*0.1, W*0.1, W*0.1])

    def prec_row_ts():
        return border_ts()

    def pr(label, key, extra=''):
        t = prec_row(label, key, extra)
        t.setStyle(border_ts())
        return t

    elements.append(pr('Is the safe distance maintained?', 'safe_distance',
        f'Yes  Voltage {prec.get("safe_distance_voltage","")}  Distance {prec.get("safe_distance_dist","")}'))
    elements.append(pr('Does the work require LOTO? If yes then LOTO Certificate must be attached.', 'loto_required'))
    elements.append(pr('Does the work require access to confined spaces? If yes, obtain a confined space entry permit', 'confined_space'))
    elements.append(pr('Have all possible sources of electrical power been isolated, locked and properly tagged (LOTO)?', 'power_isolated'))

    iso_type_row = Table([[
        P(f'{CB(prec.get("isolation_type_switch",False))} Switch Out   {CB(prec.get("isolation_type_loto",False))} Lockout/ Tag out   No. of Locks: {prec.get("num_locks","")}', 8),
        P(CB(prec.get('loto_required','')=='Yes'), 9, align=TA_CENTER),
        P('', 9, align=TA_CENTER),
        P('', 9, align=TA_CENTER),
    ]], colWidths=[W*0.7, W*0.1, W*0.1, W*0.1])
    iso_type_row.setStyle(border_ts())
    elements.append(iso_type_row)

    elements.append(pr('Has it been confirmed by testing, that the lines / equipment are de-energized ?', 'lines_de_energized'))
    elements.append(pr('Have tools and devices to be used been tested and adjusted?', 'tools_tested'))

    other_prec_row = Table([[
        P(f'Other (specify): {prec.get("other_precaution","")}', 8), P('','9'), P('','9'), P('','9'),
    ]], colWidths=[W*0.7, W*0.1, W*0.1, W*0.1])
    other_prec_row.setStyle(border_ts())
    elements.append(other_prec_row)

    # ── INSPECTED AREAS ───────────────────────────────────────────────────────
    insp_hdr = Table([[P('The following areas / items have been inspected by the issuer and receiver', 8, bold=True)]],
        colWidths=[W])
    insp_hdr.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),GRAY)]))
    elements.append(insp_hdr)

    def _i(k): return CB(insp.get(k, False))
    insp_r1 = Table([[
        P('Fire Extinguisher', 8, bold=True),
        P(f'Type {insp.get("fire_ext_type","")}', 8),
        P(f'Quantity {insp.get("fire_ext_qty","")}', 8),
        P(f'Size {insp.get("fire_ext_size","")}', 8),
    ]], colWidths=[W*0.22, W*0.26, W*0.26, W*0.26])
    insp_r1.setStyle(border_ts([('BACKGROUND',(0,0),(0,0),GRAY)]))
    elements.append(insp_r1)

    insp_r2 = Table([[
        P(f'{_i("access_escape")} Access/Escape Route', 8),
        P(f'{_i("danger_sign")} Danger/Waning Sign', 8),
        P(f'{_i("lighting")} Lighting', 8),
        P(f'{_i("safety_barriers")} Safety Barriers', 8),
    ]], colWidths=[W/4]*4)
    insp_r2.setStyle(border_ts())
    elements.append(insp_r2)

    insp_r3 = Table([[
        P(f'{_i("stick")} Stick', 8),
        P(f'{_i("portable_radio")} Portable Radio', 8),
        P(f'Other (specify): {insp.get("other_area","")}', 8),
        P('', 8),
    ]], colWidths=[W/4]*4)
    insp_r3.setStyle(border_ts())
    elements.append(insp_r3)

    # ── PPE ───────────────────────────────────────────────────────────────────
    ppe_hdr = Table([[P('PPE Required for the activity', 8, bold=True)]],
        colWidths=[W])
    ppe_hdr.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),GRAY)]))
    elements.append(ppe_hdr)

    def _p(k): return CB(ppe_d.get(k, False))
    ppe_r1 = Table([[
        P(f'{_p("helmet")} Helmet',8), P(f'{_p("safety_shoes")} Safety Shoes',8),
        P(f'{_p("elec_gloves")} Electrical Gloves  Rating {ppe_d.get("elec_gloves_rating","")}',8),
        P(f'{_p("half_mask")} Half Mask',8),
    ]], colWidths=[W*0.15, W*0.2, W*0.4, W*0.25])
    ppe_r1.setStyle(border_ts())
    elements.append(ppe_r1)

    ppe_r2 = Table([[
        P(f'{_p("safety_goggles")} Safety goggles',8),
        P(f'{_p("reflective_vest")} Reflective Vest',8),
        P(f'{_p("dust_mask")} Dust Mask',8),
        P(f'{_p("safety_clothes")} Safety clothes',8),
    ]], colWidths=[W/4]*4)
    ppe_r2.setStyle(border_ts())
    elements.append(ppe_r2)

    ppe_r3 = Table([[
        P(f'{_p("face_shield")} Face shield',8),
        P(f'{_p("arc_flash_ppe")} Arc flash PPE',8),
        P(f'{_p("ear_plugs")} Safety Ear Plugs/muff',8),
        P(f'Other: {ppe_d.get("other_ppe","")}',8),
    ]], colWidths=[W/4]*4)
    ppe_r3.setStyle(border_ts())
    elements.append(ppe_r3)

    # ── RECEIVER ACCEPTANCE ───────────────────────────────────────────────────
    issue_hdr = Table([[P('Issue and acceptance before work', 8, bold=True)]],
        colWidths=[W])
    issue_hdr.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),GRAY)]))
    elements.append(issue_hdr)

    recv_hdr = Table([[P('Acceptance of Work Permission by the person in-charge (Permit Receiver)', 8, bold=True, color=colors.HexColor('#00008b'))]],
        colWidths=[W])
    recv_hdr.setStyle(border_ts())
    elements.append(recv_hdr)

    recv_text = Table([[P(
        'I certify that I have read and verified this work permit and checklist. I have been informed about the risk assessment results. '
        'I am aware of the risks that can be exposed to. I commit that I will be in line with all the safety rules mentioned in the '
        'work permit checklist and will not deflect any of them.', 7)]], colWidths=[W])
    recv_text.setStyle(border_ts())
    elements.append(recv_text)

    recv_sig = Table([[
        P('Permit Receiver Name:', 8, bold=True), P(p.get('receiver_name',''), 9),
        P('Signature', 8, bold=True), sig_img(p.get('receiver_signature')),
        P('Date:', 8, bold=True), P(p.get('created_at','')[:10], 9),
    ]], colWidths=[W*0.18, W*0.22, W*0.1, W*0.22, W*0.08, W*0.2])
    recv_sig.setStyle(border_ts([
        ('BACKGROUND',(0,0),(0,0),GRAY),('BACKGROUND',(2,0),(2,0),GRAY),('BACKGROUND',(4,0),(4,0),GRAY),
        ('ROWHEIGHT',(0,0),(-1,-1), 16*mm),
    ]))
    elements.append(recv_sig)

    # ═══════════════════════════════════════════════════════════════════════
    # PAGE 2
    # ═══════════════════════════════════════════════════════════════════════
    elements.append(PageBreak())

    # ── ISSUER AUTHORITY ──────────────────────────────────────────────────────
    iss_hdr = Table([[P('Authority to proceed by authorized person Lead Shift Engineer (Permit Issuer)', 8, bold=True, color=colors.HexColor('#00008b'))]],
        colWidths=[W])
    iss_hdr.setStyle(border_ts())
    elements.append(iss_hdr)

    iss_text = Table([[P(
        'I reviewed the work permission checklist and checked the working conditions. I have reviewed all aspects of the '
        'task/activity and am satisfied with the arrangements as detailed in the "risk assessment" have been put in place '
        'and certify that the activity detailed above is authorized to proceed', 7)]], colWidths=[W])
    iss_text.setStyle(border_ts())
    elements.append(iss_text)

    iss_sig = Table([[
        P('Permit Issuer Name:', 8, bold=True), P(p.get('issuer_name',''), 9),
        P('Signature', 8, bold=True), sig_img(p.get('issuer_signature')),
        P('Date:', 8, bold=True), P(p.get('issued_at','')[:10] if p.get('issued_at') else '', 9),
    ]], colWidths=[W*0.18, W*0.22, W*0.1, W*0.22, W*0.08, W*0.2])
    iss_sig.setStyle(border_ts([
        ('BACKGROUND',(0,0),(0,0),GRAY),('BACKGROUND',(2,0),(2,0),GRAY),('BACKGROUND',(4,0),(4,0),GRAY),
        ('ROWHEIGHT',(0,0),(-1,-1), 16*mm),
    ]))
    elements.append(iss_sig)

    # ── HSE ENDORSEMENT ───────────────────────────────────────────────────────
    hse_hdr = Table([[P('HSE Endorsement by HSE Practitioner (O&M/EPC/Sub-Con):', 8, bold=True, color=colors.HexColor('#00008b'))]],
        colWidths=[W])
    hse_hdr.setStyle(border_ts())
    elements.append(hse_hdr)

    hse_sig = Table([[
        P('Name:', 8, bold=True), P(p.get('hse_name',''), 9),
        P('Signature', 8, bold=True), sig_img(p.get('hse_signature')),
        P('Date:', 8, bold=True), P(p.get('hse_signed_at','')[:10] if p.get('hse_signed_at') else '', 9),
    ]], colWidths=[W*0.1, W*0.28, W*0.1, W*0.24, W*0.08, W*0.2])
    hse_sig.setStyle(border_ts([
        ('BACKGROUND',(0,0),(0,0),GRAY),('BACKGROUND',(2,0),(2,0),GRAY),('BACKGROUND',(4,0),(4,0),GRAY),
        ('ROWHEIGHT',(0,0),(-1,-1), 16*mm),
    ]))
    elements.append(hse_sig)

    # ── ADDITIONAL PRECAUTIONS ────────────────────────────────────────────────
    ap_hdr = Table([[P('List of additional precaution measures if required', 8, bold=True)]],
        colWidths=[W])
    ap_hdr.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),GRAY)]))
    elements.append(ap_hdr)

    padded = (list(add_prec) + ['','','','',''])[:5]
    for i, ap in enumerate(padded, 1):
        ap_row = Table([[P(f'{i}.  {ap}', 8)]], colWidths=[W])
        ap_row.setStyle(border_ts([('MINROWHEIGHT',(0,0),(0,0), 7*mm)]))
        elements.append(ap_row)

    validity_row = Table([[P('This permit is valid for one day (one shift) from the date of issue.', 8, bold=True)]],
        colWidths=[W])
    validity_row.setStyle(border_ts())
    elements.append(validity_row)

    # ── WORKERS SIGN-IN TABLE ─────────────────────────────────────────────────
    wk_hdr_row = [
        P('Sr',8,bold=True,align=TA_CENTER), P('Name',8,bold=True), P('Iqama Number',8,bold=True),
        P('Signature',8,bold=True,align=TA_CENTER),
        P('Sr',8,bold=True,align=TA_CENTER), P('Name',8,bold=True), P('Iqama Number',8,bold=True),
        P('Signature',8,bold=True,align=TA_CENTER),
    ]
    wk_data = [wk_hdr_row]
    cw = [W*0.04, W*0.16, W*0.13, W*0.17, W*0.04, W*0.16, W*0.13, W*0.17]

    # pad workers_list to 20 slots
    wl = list(workers_list) + [{'name':'','iqama':''}] * 20
    for row_i in range(10):
        left  = wl[row_i]
        right = wl[row_i + 10]
        wk_data.append([
            P(str(row_i+1), 8, align=TA_CENTER),
            P(left.get('name',''), 8), P(left.get('iqama',''), 8), P('', 8),
            P(str(row_i+11), 8, align=TA_CENTER),
            P(right.get('name',''), 8), P(right.get('iqama',''), 8), P('', 8),
        ])

    wk_table = Table(wk_data, colWidths=cw, rowHeights=[None] + [9*mm]*10)
    wk_table.setStyle(border_ts([
        ('BACKGROUND', (0,0), (-1,0), GRAY),
        ('ROWHEIGHT', (0,1), (-1,-1), 9*mm),
    ]))
    elements.append(wk_table)

    # ── CLOSURE CHECKLIST ─────────────────────────────────────────────────────
    cl_hdr = Table([[P('CLOSURE CHECKLIST', 10, bold=True, align=TA_CENTER)]],
        colWidths=[W])
    cl_hdr.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),GRAY)]))
    elements.append(cl_hdr)

    def _c(k): return CB(closure.get(k, False))
    status_row = Table([[P('Status of Work?', 8, bold=True)]], colWidths=[W])
    status_row.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),GRAY)]))
    elements.append(status_row)

    cl_checks = Table([[
        P(f'{_c("work_completed")} Work Completed', 8),
        P(f'{_c("work_incomplete")} Work Incomplete', 8),
        P(f'{_c("tools_removed")} All Tools & Equipment Removed', 8),
    ],[
        P(f'{_c("loto_closed")} LOTO / ICC Closed, If Applied', 8),
        P(f'{_c("permit_suspended")} Permit Suspended', 8),
        P(f'{_c("housekeeping_done")} Housekeeping Done', 8),
    ]], colWidths=[W/3]*3)
    cl_checks.setStyle(border_ts())
    elements.append(cl_checks)

    # Closure signatures
    cl_sig = Table([[
        P('Permit Receiver', 8, bold=True, align=TA_CENTER),
        P('Permit Issuer', 8, bold=True, align=TA_CENTER),
    ]], colWidths=[W/2, W/2])
    cl_sig.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),GRAY)]))
    elements.append(cl_sig)

    cl_sig2 = Table([[
        P('Name:', 8), P(p.get('receiver_name',''), 8),
        P('Signature:', 8), P('', 8),
        P('Date:', 8), P('', 8),
        P('Name:', 8), P(p.get('issuer_name',''), 8),
        P('Signature:', 8), P('', 8),
        P('Date:', 8), P('', 8),
    ]], colWidths=[W*0.06, W*0.13, W*0.08, W*0.14, W*0.06, W*0.13,
                   W*0.06, W*0.13, W*0.08, W*0.14, W*0.06, W*0.13])
    cl_sig2.setStyle(border_ts([('ROWHEIGHT',(0,0),(-1,-1), 10*mm)]))
    elements.append(cl_sig2)

    hse_cl_hdr = Table([[P('HSE Practitioner (O&M/EPC/Sub-Con)', 8, bold=True, align=TA_CENTER)]],
        colWidths=[W])
    hse_cl_hdr.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),GRAY)]))
    elements.append(hse_cl_hdr)

    hse_cl_sig = Table([[
        P('Name:', 8), P(p.get('hse_name',''), 8),
        P('Signature:', 8), P('', 8),
        P('Date:', 8), P('', 8),
    ]], colWidths=[W*0.08, W*0.26, W*0.1, W*0.28, W*0.08, W*0.2])
    hse_cl_sig.setStyle(border_ts([('ROWHEIGHT',(0,0),(-1,-1), 10*mm)]))
    elements.append(hse_cl_sig)

    # ── Footer ────────────────────────────────────────────────────────────────
    elements.append(Spacer(1, 2*mm))
    ft = Table([[P(f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}  |  HDEC CMMS  |  PTW: {ptw or "PENDING"}', 7, align=TA_CENTER)]],
        colWidths=[W])
    elements.append(ft)

    doc.build(elements)
    buffer.seek(0)
    return buffer


# ── ICC PDF (Isolation Confirmation Certificate — MP-10 FORM 3) ────────────
def generate_icc_pdf(permit_id: str) -> io.BytesIO | None:
    """Generate the Isolation Confirmation Certificate matching the MP-10 ICC format."""
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        from reportlab.lib.units import mm
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Table, TableStyle,
            Spacer, PageBreak, Image as RLImage
        )
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
        import base64, tempfile
    except ImportError:
        return None

    permit = get_permit(permit_id)
    if not permit or not permit.get('isolation_required'):
        return None

    buffer = io.BytesIO()
    LM = RM = 12*mm
    W = A4[0] - LM - RM
    doc = SimpleDocTemplate(buffer, pagesize=A4,
        leftMargin=LM, rightMargin=RM, topMargin=12*mm, bottomMargin=12*mm)

    BLK  = colors.black
    GRAY = colors.HexColor('#f2f2f2')
    RED  = colors.HexColor('#c00000')
    YEL  = colors.HexColor('#ffff00')

    def P(txt, size=8, bold=False, align=TA_LEFT, color=BLK):
        fn = 'Helvetica-Bold' if bold else 'Helvetica'
        return Paragraph(str(txt or ''), ParagraphStyle('_p',
            fontName=fn, fontSize=size, textColor=color,
            alignment=align, leading=size*1.3, spaceBefore=1, spaceAfter=1))

    def CB(v): return '■' if v else '□'

    def border_ts(extra=None):
        s = [('BOX',(0,0),(-1,-1),0.5,BLK),('INNERGRID',(0,0),(-1,-1),0.5,BLK),
             ('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2),
             ('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),
             ('VALIGN',(0,0),(-1,-1),'MIDDLE')]
        if extra: s.extend(extra)
        return TableStyle(s)

    p = permit
    iso_type = p.get('isolation_type', {})
    iso_seq  = p.get('isolation_sequence', [])
    deiso_seq= p.get('de_isolation_sequence', [])
    icc_no   = p.get('isolation_cert_number', '')
    ptw_no   = p.get('permit_number', '')
    valid_from = str(p.get('valid_from',''))

    elements = []

    # ── HEADER ────────────────────────────────────────────────────────────────
    hdr = Table([[
        P('POWERCHINA', 10, bold=True),
        P('Isolation Confirmation Certificate\nMP-10 FORM 3 Isolation Confirmation Certificate', 11, bold=True, align=TA_CENTER, color=RED),
        P('SANA TAIBAH', 10, bold=True, align=TA_CENTER),
    ]], colWidths=[W*0.2, W*0.6, W*0.2])
    hdr.setStyle(border_ts())
    elements.append(hdr)

    sub = Table([[P('1100MW Al Henakiyah Solar Photovoltaic Independent Power Plant', 9, bold=True, align=TA_CENTER)]],
        colWidths=[W])
    sub.setStyle(border_ts())
    elements.append(sub)

    # ── REF ROW ───────────────────────────────────────────────────────────────
    ref1 = Table([[
        P('Isolation Confirmation\nCertificate (ICC) No. :', 8, bold=True, color=RED),
        P(icc_no, 9),
        P('Date :', 8, bold=True), P(valid_from[:10], 9),
        P('Time:', 8, bold=True), P(valid_from[11:16] if len(valid_from)>10 else '', 9),
        P('Additional LOTO List:', 8, bold=True), P('', 9),
    ]], colWidths=[W*0.18, W*0.16, W*0.06, W*0.1, W*0.06, W*0.1, W*0.14, W*0.2])
    ref1.setStyle(border_ts([('BACKGROUND',(0,0),(0,0),GRAY),('BACKGROUND',(2,0),(2,0),GRAY),
                              ('BACKGROUND',(4,0),(4,0),GRAY),('BACKGROUND',(6,0),(6,0),GRAY)]))
    elements.append(ref1)

    ref2 = Table([[
        P('PTW No. :', 8, bold=True, color=RED), P(ptw_no, 9),
        P('SLD/ Drawing No.:', 8, bold=True), P(p.get('sld_drawing_no',''), 9),
    ]], colWidths=[W*0.1, W*0.24, W*0.16, W*0.5])
    ref2.setStyle(border_ts([('BACKGROUND',(0,0),(0,0),GRAY),('BACKGROUND',(2,0),(2,0),GRAY)]))
    elements.append(ref2)

    # ── PERMIT RECEIVER DETAILS + DESCRIPTION ─────────────────────────────────
    detail = Table([
        [P('Permit Receiver  Details', 8, bold=True),
         P('Name:', 8, bold=True), P(p.get('receiver_name',''), 8),
         P('Description /Work To be Performed:', 8, bold=True)],
        ['', P('Company:', 8, bold=True), P(p.get('receiver_company',''), 8),
         P(p.get('job_description',''), 8)],
        ['', P('ID Number:', 8, bold=True), P(p.get('receiver_id',''), 8), ''],
        ['', P('Signature', 8, bold=True), P('', 8), ''],
    ], colWidths=[W*0.18, W*0.12, W*0.24, W*0.46])
    detail.setStyle(border_ts([
        ('SPAN',(0,0),(0,3)), ('SPAN',(3,0),(3,3)),
        ('BACKGROUND',(0,0),(0,3),GRAY),
        ('ROWHEIGHT',(0,3),(-1,3), 14*mm),
    ]))
    elements.append(detail)

    # ── ISOLATION TYPE ─────────────────────────────────────────────────────────
    it = iso_type
    iso_type_t = Table([[
        P(f'Isolation Type (Plz Tick all that Apply):', 8, bold=True, color=RED),
        P(f'{CB(it.get("electrical",False))} Electrical Isolation', 8),
        P(f'{CB(it.get("mechanical",False))} Mechanical Isolation', 8),
        P(f'{CB(it.get("earthing",False))} Earthing', 8),
        P(f'Others (Plz Specify): {it.get("others","")}', 8),
    ]], colWidths=[W*0.25, W*0.18, W*0.18, W*0.14, W*0.25])
    iso_type_t.setStyle(border_ts())
    elements.append(iso_type_t)

    # ── ISOLATION SECTION HEADER ───────────────────────────────────────────────
    iso_hdr_t = Table([[P('ISOLATION', 10, bold=True, align=TA_CENTER)]],
        colWidths=[W])
    iso_hdr_t.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),YEL)]))
    elements.append(iso_hdr_t)

    receiver_decl = Table([[P(
        'I, as Authorized Permit Reciever, declare that the plant/system detailed in this ICC request is in a safe '
        'condition for isolations to be commenced. I request that the actions for isolation shall be taken in the '
        'following sequence:', 7)]], colWidths=[W])
    receiver_decl.setStyle(border_ts())
    elements.append(receiver_decl)

    # ── ISOLATION SEQUENCE TABLE ───────────────────────────────────────────────
    iso_seq_hdr = [
        [P('ISOLATION SEQUENCE (As per Single Line Diagram / Switching Program)', 8, bold=True, align=TA_CENTER),
         '', '', '', '', '', '', '', '', '', ''],
    ]
    iso_seq_hdr2 = [
        P('Device/\nEquipment\nIsolated',7,bold=True,align=TA_CENTER),
        P('Isolation\nPoint',7,bold=True,align=TA_CENTER),
        P('Location',7,bold=True,align=TA_CENTER),
        P('Applied\nTag No.',7,bold=True,align=TA_CENTER),
        P('Lock\nApplied\nKey No.',7,bold=True,align=TA_CENTER),
        P('Actions\nOff /\nOpen',7,bold=True,align=TA_CENTER),
        P('On /\nClosed',7,bold=True,align=TA_CENTER),
        P('Earthing\nApplied\nLocation',7,bold=True,align=TA_CENTER),
        P('EQPT/\nES Point',7,bold=True,align=TA_CENTER),
        P('Date',7,bold=True,align=TA_CENTER),
        P('Time',7,bold=True,align=TA_CENTER),
        P('Name and\nIqama Number',7,bold=True,align=TA_CENTER),
        P('Signature',7,bold=True,align=TA_CENTER),
    ]
    iso_cw = [W*0.1, W*0.09, W*0.09, W*0.07, W*0.07, W*0.05, W*0.05,
              W*0.09, W*0.07, W*0.06, W*0.05, W*0.11, W*0.1]

    padded_seq = (list(iso_seq) + [{}]*6)[:6]
    iso_rows = [iso_seq_hdr2]
    for row in padded_seq:
        iso_rows.append([
            P(row.get('device',''),7), P(row.get('iso_point',''),7),
            P(row.get('location',''),7), P(row.get('tag_no',''),7),
            P(row.get('lock_key',''),7), P(row.get('off_open',''),7),
            P(row.get('on_closed',''),7), P(row.get('earth_location',''),7),
            P(row.get('eqpt_es',''),7), P(row.get('date',''),7),
            P(row.get('time',''),7), P(row.get('name_iqama',''),7), P('',7),
        ])

    iso_seq_t = Table(iso_rows, colWidths=iso_cw, rowHeights=[None]+[10*mm]*6)
    iso_seq_t.setStyle(border_ts([('BACKGROUND',(0,0),(-1,0),GRAY)]))
    elements.append(iso_seq_t)

    # ── APPROVAL SIGNATURES ────────────────────────────────────────────────────
    def sig_row(role_label, name='', sig_b64=None, iqama='', date_str=''):
        return Table([[
            P(role_label, 7, bold=True),
            P('Name', 7, bold=True), P(name, 7),
            P('Signature', 7, bold=True), P('' if not sig_b64 else '(signed)', 7),
            P('Iqama Number', 7, bold=True), P(iqama, 7),
            P('Date and Time', 7, bold=True), P(date_str, 7),
        ]], colWidths=[W*0.22, W*0.06, W*0.14, W*0.08, W*0.14, W*0.1, W*0.1, W*0.08, W*0.08])

    def srow(label, name='', date=''):
        t = Table([[
            P(label, 7, bold=True, color=colors.HexColor('#00008b')),
            P('Name', 7, bold=True), P(name, 7),
            P('Signature', 7, bold=True), P('', 7),
            P('Iqama Number', 7, bold=True), P('', 7),
            P('Date and Time', 7, bold=True), P(date, 7),
        ]], colWidths=[W*0.24, W*0.06, W*0.12, W*0.08, W*0.14, W*0.1, W*0.1, W*0.08, W*0.08])
        t.setStyle(border_ts([('ROWHEIGHT',(0,0),(-1,-1),10*mm)]))
        return t

    elements.append(srow('Isolation Approved & verified by\n(Plant Manager/Operation Manager) :',
        p.get('issuer_name',''), p.get('issued_at','')[:10] if p.get('issued_at') else ''))
    elements.append(srow('Isolation accepted & Verified by\n(Authorized PTW Receiver) :',
        p.get('receiver_name',''), p.get('created_at','')[:10]))
    elements.append(srow('Isolation Done by\n(Lead Shift Engineer) :'))
    elements.append(srow('Isolation & HSE Checks Verified by\n(HSE Practitioner O&M/Sub-Con) :',
        p.get('hse_name',''), p.get('hse_signed_at','')[:10] if p.get('hse_signed_at') else ''))

    # ═══════════════════════════════════════════════════════════════════════
    # PAGE 2 — DE-ISOLATION
    # ═══════════════════════════════════════════════════════════════════════
    elements.append(PageBreak())

    deiso_hdr = Table([[P('DE-ISOLATION SEQUENCE (As per Single Line Diagram/ Switching Program)', 8, bold=True, align=TA_CENTER)]],
        colWidths=[W])
    deiso_hdr.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),GRAY)]))
    elements.append(deiso_hdr)

    # De-isolation sequence table (same columns but "removed" instead of "applied")
    deiso_hdr2 = [
        P('Device/\nEquipment\nDe-Isolated',7,bold=True,align=TA_CENTER),
        P('De-Isolation\nPoint',7,bold=True,align=TA_CENTER),
        P('Location',7,bold=True,align=TA_CENTER),
        P('Removed\nTag No.',7,bold=True,align=TA_CENTER),
        P('Lock\nRemoved\nKey No.',7,bold=True,align=TA_CENTER),
        P('Actions\nOff /\nOpen',7,bold=True,align=TA_CENTER),
        P('On /\nClosed',7,bold=True,align=TA_CENTER),
        P('Earthing\nRemoved\nLocation',7,bold=True,align=TA_CENTER),
        P('EQPT/\nES Point',7,bold=True,align=TA_CENTER),
        P('Date',7,bold=True,align=TA_CENTER),
        P('Time',7,bold=True,align=TA_CENTER),
        P('Name and\nIqama Number',7,bold=True,align=TA_CENTER),
        P('Signature',7,bold=True,align=TA_CENTER),
    ]
    padded_deiso = (list(deiso_seq) + [{}]*6)[:6]
    deiso_rows = [deiso_hdr2]
    for row in padded_deiso:
        deiso_rows.append([
            P(row.get('device',''),7), P(row.get('iso_point',''),7),
            P(row.get('location',''),7), P(row.get('tag_no',''),7),
            P(row.get('lock_key',''),7), P(row.get('off_open',''),7),
            P(row.get('on_closed',''),7), P(row.get('earth_location',''),7),
            P(row.get('eqpt_es',''),7), P(row.get('date',''),7),
            P(row.get('time',''),7), P(row.get('name_iqama',''),7), P('',7),
        ])
    deiso_t = Table(deiso_rows, colWidths=iso_cw, rowHeights=[None]+[10*mm]*6)
    deiso_t.setStyle(border_ts([('BACKGROUND',(0,0),(-1,0),GRAY)]))
    elements.append(deiso_t)

    # Precautions taken
    prec_taken = Table([
        [P('Precautions Taken :', 8, bold=True)],
        [P('', 8)],
    ], colWidths=[W])
    prec_taken.setStyle(border_ts([('MINROWHEIGHT',(0,1),(0,1), 10*mm)]))
    elements.append(prec_taken)

    # Work completion & de-isolation signatures
    completion_decl = Table([[P(
        'I hereby confirm that all work has been completed, housekeeping is done, Workers and materials are removed '
        'from the site/plant/system specified on this isolation confirmation certificate and operational integrity '
        'precautions have been taken, now the normal operations may safely be reinstated.', 7)]], colWidths=[W])
    completion_decl.setStyle(border_ts([('BACKGROUND',(0,0),(-1,-1),YEL)]))
    elements.append(completion_decl)

    elements.append(srow('Work Completion Confirmed by (Authorized PTW Receiver):',
        p.get('receiver_name',''), ''))
    elements.append(Table([[P('Removed LOTO devices and restored system as per approved de-isolation procedure.', 7)]],
        colWidths=[W]))
    elements.append(srow('De-ioslation Carried out by (Lead Shift Engineer):'))
    elements.append(Table([[P('Verified safe restoration of equipment and confirmed system integrity.', 7)]],
        colWidths=[W]))
    elements.append(srow('De-ioslation Authorized by (Plant Manager/Operation Manager):'))
    elements.append(Table([[P('Verified that de-isolation and system restoration were conducted safely and in compliance with HSE requirements.', 7)]],
        colWidths=[W]))
    elements.append(srow('De-ioslation verified by (O&M HSE Practitioner):'))

    doc.build(elements)
    buffer.seek(0)
    return buffer


# ── WORD (.docx) PERMIT EXPORT ────────────────────────────────────────────

def _b64_to_stream(data_url: str):
    """Convert a base64 data URL (data:image/...;base64,...) to a BytesIO stream."""
    import base64 as _b64
    if not data_url:
        return None
    try:
        header, data = data_url.split(',', 1)
        return io.BytesIO(_b64.b64decode(data))
    except Exception:
        return None


def _add_sig_image_to_para(para, sig_b64: str, max_w_cm: float, max_h_cm: float):
    """
    Add a signature image run to *para*, scaled so it fits within
    max_w_cm × max_h_cm while preserving the image aspect ratio.
    Uses PIL to read the real pixel dimensions.
    """
    stream = _b64_to_stream(sig_b64)
    if not stream:
        return
    try:
        from PIL import Image as _PilImg
        from docx.shared import Cm
        stream.seek(0)
        img = _PilImg.open(stream)
        img_w, img_h = img.size
        aspect = (img_h / img_w) if img_w > 0 else 1.0

        # Fit to max width first, then clamp to max height
        final_w = max_w_cm
        final_h = final_w * aspect
        if final_h > max_h_cm:
            final_h = max_h_cm
            final_w = final_h / aspect if aspect > 0 else final_w

        stream.seek(0)
        run = para.add_run()
        run.add_picture(stream, width=Cm(final_w), height=Cm(final_h))
    except Exception:
        pass


def _set_para_text(para, text: str, bold: bool = False, font_size_pt: float = 9):
    """Replace all runs in *para* with a single run containing *text*."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
    for r_el in list(para._p.findall(qn('w:r'))):
        para._p.remove(r_el)
    r_new = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    if bold:
        b_el = OxmlElement('w:b')
        rpr.append(b_el)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(font_size_pt * 2)))
    rpr.append(sz)
    r_new.append(rpr)
    t_new = OxmlElement('w:t')
    t_new.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t_new.text = text or ''
    r_new.append(t_new)
    para._p.append(r_new)


def _fill_t7_sig_cell(cell, role_label: str, name: str, date_str: str, sig_b64: str):
    """
    Fill a Table-7 (Issue & Acceptance / Work-Started) signature cell.

    Template structure inside each cell:
      Para[0] — role title (keep as-is)
      Para[1] — declaration text (keep as-is)
      Para[2] — "Permit Receiver Name: ...  Signature  ...  Date:" — REPLACE

    After Para[2] we insert:
      • a paragraph with the e-signature image (max 5.5 cm × 1.8 cm)
      • a paragraph with "Date: <date_str>"
    """
    paras = cell.paragraphs
    # Rewrite Para[2]: show "Permit Receiver Name: [name]  |  Approved by: [name]"
    if len(paras) > 2:
        _set_para_text(paras[2], f'{role_label}Name: {name or ""}', font_size_pt=9)

    # Append signature image
    if sig_b64:
        sig_para = cell.add_paragraph()
        _add_sig_image_to_para(sig_para, sig_b64, max_w_cm=5.5, max_h_cm=1.8)

    # Append date
    date_para = cell.add_paragraph()
    _set_para_text(date_para, f'Date: {date_str or ""}', font_size_pt=9)


def _fill_t10_name_cell(cell, name: str, date: str = None):
    """
    Fill a closure table Name cell (Table 10, rows 3/5, cols 0/3/0).
    Replaces the template 'Name:' text with the actual name.
    Optionally sets date if *date* is given (for single-cell layouts).
    """
    if cell.paragraphs:
        _set_para_text(cell.paragraphs[0], f'Name: {name or ""}', font_size_pt=9)
    if date and len(cell.paragraphs) > 1:
        _set_para_text(cell.paragraphs[1], f'Date: {date or ""}', font_size_pt=9)


def _fill_t10_date_cell(cell, date_str: str):
    """Fill a closure table Date cell (Table 10, rows 3/5, cols 2/5/2)."""
    if cell.paragraphs:
        _set_para_text(cell.paragraphs[0], f'Date: {date_str or ""}', font_size_pt=9)


def _fill_t10_sig_cell(cell, sig_b64: str, cell_width_cm: float):
    """
    Fill a closure table Signature cell (Table 10, rows 3/5, cols 1/4/1).
    Clears the 'Signature:' label and inserts the image, fitted to the cell.
    Keeps a small margin: max_w = cell_width - 0.3 cm, max_h = 1.8 cm.
    """
    if cell.paragraphs:
        _set_para_text(cell.paragraphs[0], '', font_size_pt=9)
    if sig_b64:
        max_w = max(cell_width_cm - 0.3, 0.5)
        sig_para = cell.add_paragraph()
        _add_sig_image_to_para(sig_para, sig_b64, max_w_cm=max_w, max_h_cm=1.8)


# Keep for backwards-compat (used elsewhere in the file)
def _insert_sig_image_in_cell(cell, sig_b64: str, width_cm: float = 3.5):
    stream = _b64_to_stream(sig_b64)
    if not stream:
        return
    try:
        from docx.shared import Cm
        para = cell.add_paragraph()
        run = para.add_run()
        run.add_picture(stream, width=Cm(width_cm))
    except Exception:
        pass


def generate_permit_docx(permit_id: str) -> io.BytesIO:
    """
    Fill the MP-10 FORM 3 General Work Permit DOCX template with permit data
    and return as an in-memory BytesIO buffer.
    """
    from docx import Document as DocxDocument
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    permit = get_permit(permit_id)
    if not permit:
        raise ValueError(f'Permit {permit_id} not found')

    template_path = MEDIA_ROOT / 'MP-10 FORM 3 General Work Permit.docx'
    if not template_path.exists():
        raise FileNotFoundError(f'Template not found: {template_path}')

    doc = DocxDocument(str(template_path))

    # ── Helpers ────────────────────────────────────────────────────────────
    def set_cell_text(cell, text: str):
        """Clear all runs from first paragraph, add plain text run."""
        para = cell.paragraphs[0]
        p_el = para._p
        for r_el in list(p_el.findall(qn('w:r'))):
            p_el.remove(r_el)
        r_new = OxmlElement('w:r')
        t_new = OxmlElement('w:t')
        t_new.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t_new.text = text or ''
        r_new.append(t_new)
        p_el.append(r_new)

    def mark_checkbox_cell(cell, checked: bool):
        """Replace inline image checkbox run with ☑/☐ + label text."""
        mark = '☑  ' if checked else '☐  '
        for para in cell.paragraphs:
            p_el = para._p
            # Collect label text from non-drawing runs
            label_parts = []
            for r_el in list(p_el.findall(qn('w:r'))):
                if r_el.find(qn('w:drawing')) is not None:
                    p_el.remove(r_el)
                else:
                    t = r_el.find(qn('w:t'))
                    if t is not None:
                        label_parts.append(t.text or '')
                    p_el.remove(r_el)
            label = ''.join(label_parts).strip()
            # Insert single run with mark + label
            r_new = OxmlElement('w:r')
            rpr = OxmlElement('w:rPr')
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), '14')
            rpr.append(sz)
            r_new.append(rpr)
            t_new = OxmlElement('w:t')
            t_new.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t_new.text = f'{mark}{label}'
            r_new.append(t_new)
            p_el.append(r_new)
            break  # only first paragraph

    # ── Permit data shortcuts ──────────────────────────────────────────────
    p = permit
    risks        = set(p.get('risks', []))
    precautions  = set(p.get('precaution_checks', []))
    ppe_set      = set(p.get('ppe_required', []))
    attached_docs = set(p.get('attached_docs', []))
    workers      = p.get('workers_list', [])
    closure      = p.get('closure', {})

    # ── TABLE 0: Header ────────────────────────────────────────────────────
    t0 = doc.tables[0]
    set_cell_text(t0.rows[0].cells[1], p.get('permit_no', p.get('id', '')[:8].upper()))
    set_cell_text(t0.rows[0].cells[3], p.get('date', p.get('created_at', '')[:10]))
    set_cell_text(t0.rows[1].cells[3], p.get('company', 'HDEC'))
    set_cell_text(t0.rows[2].cells[1], p.get('area', p.get('location', '')))
    set_cell_text(t0.rows[2].cells[3], p.get('department', 'Maintenance'))

    # ── TABLE 1: Description of Work ──────────────────────────────────────
    t1 = doc.tables[1]
    desc = p.get('work_description', '') or p.get('description', '')
    if not desc and p.get('equipment'):
        desc = p['equipment']
    set_cell_text(t1.rows[1].cells[0], desc)

    # ── TABLE 2: Tools and Equipment ──────────────────────────────────────
    t2 = doc.tables[2]
    set_cell_text(t2.rows[1].cells[0], p.get('tools_equipment', p.get('tools', '')))

    # ── TABLE 3: Hazards ──────────────────────────────────────────────────
    HAZARD_MAP = {
        'electric arc flash': 'electric_arc_flash',
        'contact to hv/mv/lv/live working': 'hv_mv_lv',
        'confined space': 'confined_space',
        'manual handling': 'manual_handling',
        'fall from height / same level': 'fall_from_height',
        'falling objects': 'falling_objects',
        'impact load / objects': 'impact_load',
        'high temperature': 'high_temperature',
        'corrosive substances': 'corrosive_substances',
        'stored energy': 'stored_energy',
        'pressurized system': 'pressurized_system',
        'heat stress': 'heat_stress',
        'improper storage': 'improper_storage',
        'over-exertion': 'over_exertion',
        'entrapment / entanglement': 'entrapment',
        'lifting / overhead loads': 'lifting_loads',
        'radiation': 'radiation',
        'physical agent': 'physical_agent',
        'ergonomics': 'ergonomics',
        'flammable material': 'flammable_material',
        'dust exposures': 'dust_exposures',
        'hot surfaces': 'hot_surfaces',
        'toxic agents': 'toxic_agents',
        'access / egress': 'access_egress',
        'cut / laceration': 'cut_laceration',
        'housekeeping': 'housekeeping',
        'asphyxiation': 'asphyxiation',
        'un-insulated tools': 'uninsulated_tools',
        'lone working': 'lone_working',
        'flammable / explosives': 'flammable_explosives',
        'dust / fumes': 'dust_fumes',
        'slips, trips & falls': 'slips_trips',
        'underground services (cables etc.)': 'underground_services',
        'high noise': 'high_noise',
        'others': 'hazards_other',
    }
    t3 = doc.tables[3]
    for row in t3.rows[2:]:
        for cell in row.cells:
            raw = cell.text.strip().lower()
            if not raw:
                continue
            key = HAZARD_MAP.get(raw, raw.replace(' ', '_').replace('/', '_'))
            checked = key in risks or raw in {r.lower() for r in risks}
            mark_checkbox_cell(cell, checked)

    # ── TABLE 4: Precautions ───────────────────────────────────────────────
    PRECAUTION_MAP = {
        'equipment locked': 'equipment_locked',
        'earthing installed': 'earthing_installed',
        'cable discharged': 'cable_discharged',
        'fuses withdrawn': 'fuses_withdrawn',
        'jsa / safe work system attached': 'jsa_attached',
        'ra / ms attached': 'ra_ms_attached',
        'fire extinguisher / fire blanket': 'fire_extinguisher',
        'first aid available': 'first_aid',
        'insulated tools available': 'insulated_tools',
        'lock out tags applied': 'lockout_tags',
        'forced ventilation': 'forced_ventilation',
        'proper access & egress': 'proper_access',
        'pre-job meeting / toolbox talks': 'toolbox_talks',
        'area barricades / guards available': 'area_barricades',
        'warning signs displayed': 'warning_signs',
        'isolation tested & verified': 'isolation_tested',
        'others': 'precaution_other',
    }
    t4 = doc.tables[4]
    for row in t4.rows[2:]:
        for cell in row.cells:
            raw = cell.text.strip()
            if not raw:
                continue
            # Cells can have tab-separated sub-items
            parts = [x.strip() for x in raw.replace('\t', '\n').split('\n') if x.strip()]
            any_checked = any(
                PRECAUTION_MAP.get(part.lower(), part.lower().replace(' ', '_')) in precautions
                or part.lower() in {x.lower() for x in precautions}
                for part in parts
            )
            mark_checkbox_cell(cell, any_checked)

    # ── TABLE 5: PPE ───────────────────────────────────────────────────────
    PPE_MAP = {
        'hard helmet': 'hard_helmet',
        'safety glasses / goggles': 'safety_glasses',
        'safety shoes': 'safety_shoes',
        'leather / electrical insulated gloves': 'insulated_gloves',
        'welding shield': 'welding_shield',
        'respirators / dust masks': 'dust_mask',
        'work gloves': 'work_gloves',
        'ear plugs / muffs': 'ear_protection',
        'coveralls / work clothes': 'coveralls',
        'chemical gloves': 'chemical_gloves',
        'face shield': 'face_shield',
        'others': 'ppe_other',
    }
    t5 = doc.tables[5]
    for row in t5.rows[2:]:
        for cell in row.cells:
            raw = cell.text.strip()
            if not raw:
                continue
            key = PPE_MAP.get(raw.lower(), raw.lower().replace(' ', '_').replace('/', '_'))
            checked = key in ppe_set or raw.lower() in {x.lower() for x in ppe_set}
            mark_checkbox_cell(cell, checked)

    # ── TABLE 6: Attached Documents ────────────────────────────────────────
    DOCS_MAP = {
        'drawings / layouts': 'drawings',
        'hot work permit': 'hot_work_permit',
        'confined space entry permit': 'confined_space_permit',
        'loto permit / isolation certificate': 'loto_permit',
        'work at height permit': 'height_permit',
        'lifting work permit': 'lifting_permit',
        'emergency response plan': 'emergency_plan',
        'personnel qualification certificates': 'qualifications',
        'excavation work permit': 'excavation_permit',
        'approved risk assessment /method': 'risk_assessment',
        'authorized personnel list   statement': 'authorized_list',
        'others': 'docs_other',
    }
    t6 = doc.tables[6]
    for row in t6.rows[2:]:
        for cell in row.cells:
            raw = cell.text.strip()
            if not raw:
                continue
            parts = [x.strip() for x in raw.replace('\t', '\n').split('\n') if x.strip()]
            any_checked = any(
                DOCS_MAP.get(part.lower(), part.lower().replace(' ', '_')) in attached_docs
                or part.lower() in {x.lower() for x in attached_docs}
                for part in parts
            )
            mark_checkbox_cell(cell, any_checked)

    # ── TABLE 7: Issue & Acceptance Signatures (Work Started) ─────────────────
    # Each row is a single wide cell (6.62 in) with 3 paragraphs:
    #   Para[0] = role title,  Para[1] = declaration,  Para[2] = Name/Sig/Date line
    # We rewrite Para[2] with the actual name, then append: image + date.
    t7 = doc.tables[7]
    _fill_t7_sig_cell(
        t7.rows[1].cells[0],
        role_label='Permit Receiver ',
        name=p.get('receiver_name', ''),
        date_str=(p.get('application_datetime') or p.get('created_at', ''))[:10],
        sig_b64=p.get('receiver_signature'),
    )
    _fill_t7_sig_cell(
        t7.rows[2].cells[0],
        role_label='Permit Issuer ',
        name=p.get('issuer_name', ''),
        date_str=(p.get('issued_at', '') or '')[:10],
        sig_b64=p.get('issuer_signature'),
    )
    _fill_t7_sig_cell(
        t7.rows[3].cells[0],
        role_label='HSE Practitioner ',
        name=p.get('hse_name', ''),
        date_str=(p.get('hse_signed_at', '') or '')[:10],
        sig_b64=p.get('hse_signature'),
    )

    # ── TABLE 9: Workers ───────────────────────────────────────────────────
    t9 = doc.tables[9]
    for i, worker in enumerate(workers[:10]):
        row = t9.rows[i + 1]
        set_cell_text(row.cells[1], worker.get('name', ''))
        set_cell_text(row.cells[2], worker.get('iqama', ''))
    for i, worker in enumerate(workers[10:20]):
        row = t9.rows[i + 1]
        set_cell_text(row.cells[5], worker.get('name', ''))
        set_cell_text(row.cells[6], worker.get('iqama', ''))

    # ── TABLE 10: Closure Signatures ──────────────────────────────────────
    # Row 3 layout (6 cells):
    #   [0] Name:  [1] Signature:  [2] Date:  |  [3] Name:  [4] Signature:  [5] Date:
    #      Permit Receiver                           Permit Issuer
    # Row 5 layout (6 cells):
    #   [0] Name:  [1] Signature:  [2] Date:
    #      HSE Practitioner
    t10 = doc.tables[10]
    closed_at = (p.get('closed_at', '') or '')[:10]

    if len(t10.rows) > 3 and len(t10.rows[3].cells) >= 6:
        r3 = t10.rows[3].cells
        # Permit Receiver
        _fill_t10_name_cell(r3[0], name=p.get('receiver_name', ''))
        _fill_t10_sig_cell(r3[1], sig_b64=p.get('closure_receiver_signature'), cell_width_cm=3.04)
        _fill_t10_date_cell(r3[2], date_str=closed_at)
        # Permit Issuer
        _fill_t10_name_cell(r3[3], name=p.get('issuer_name', ''))
        _fill_t10_sig_cell(r3[4], sig_b64=p.get('closure_issuer_signature'), cell_width_cm=2.59)
        _fill_t10_date_cell(r3[5], date_str=closed_at)

    if len(t10.rows) > 5 and len(t10.rows[5].cells) >= 3:
        r5 = t10.rows[5].cells
        # HSE Practitioner
        _fill_t10_name_cell(r5[0], name=p.get('hse_name', ''))
        _fill_t10_sig_cell(r5[1], sig_b64=p.get('closure_hse_signature'), cell_width_cm=3.04)
        _fill_t10_date_cell(r5[2], date_str=closed_at)

    # ── Return as bytes ────────────────────────────────────────────────────
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ── HANDOVER / SHIFT LOG ──────────────────────────────────────────────────

def get_handovers() -> list:
    return _load(HANDOVERS_FILE)


def get_handover(handover_id: str) -> dict | None:
    return next((h for h in get_handovers() if h['id'] == handover_id), None)


def get_handovers_by_date(date_str: str) -> dict:
    """Return {'day': handover_or_None, 'night': handover_or_None} for a date."""
    all_h = [h for h in get_handovers() if h.get('date') == date_str]
    return {
        'day':   next((h for h in all_h if h.get('shift') == 'day'),   None),
        'night': next((h for h in all_h if h.get('shift') == 'night'), None),
    }


def get_handover_dates() -> list:
    """Return sorted unique dates that have handover records (newest first)."""
    dates = sorted({h['date'] for h in get_handovers() if h.get('date')}, reverse=True)
    return dates


def create_handover(data: dict) -> dict:
    handovers = get_handovers()
    h = {
        'id':                       str(uuid.uuid4()),
        'date':                     data.get('date', datetime.now().strftime('%Y-%m-%d')),
        'shift':                    data.get('shift', 'day'),   # 'day' or 'night'
        'timing':                   '08:00 – 20:00' if data.get('shift') == 'day' else '20:00 – 08:00',
        'shift_incharge':           data.get('shift_incharge', ''),
        'technicians':              data.get('technicians', ''),
        'major_alarms':             data.get('major_alarms', ''),
        'equipment_breakdown':      data.get('equipment_breakdown', ''),
        'maintenance_activities':   data.get('maintenance_activities', []),
        'inverter_faults':          data.get('inverter_faults', []),
        'scb_faults':               data.get('scb_faults', []),
        'robot_maintenance':        data.get('robot_maintenance', []),
        'spare_parts':              data.get('spare_parts', []),
        'key_issues':               data.get('key_issues', ''),
        'pending_work':             data.get('pending_work', ''),
        'instructions_next_shift':  data.get('instructions_next_shift', ''),
        'observation_text':         data.get('observation_text', ''),
        'observation_images':       [],
        'shift_engineer_sig':       data.get('shift_engineer_sig', ''),
        'incoming_engineer_sig':    data.get('incoming_engineer_sig', ''),
        'status':                   'draft',
        'created_at':               datetime.now().isoformat(),
        'created_by':               data.get('created_by', ''),
        'submitted_at':             None,
    }
    handovers.append(h)
    _save(HANDOVERS_FILE, handovers)
    return h


def update_handover(handover_id: str, updates: dict) -> dict | None:
    handovers = get_handovers()
    for i, h in enumerate(handovers):
        if h['id'] == handover_id:
            handovers[i].update(updates)
            _save(HANDOVERS_FILE, handovers)
            return handovers[i]
    return None


def save_handover_image(handover_id: str, uploaded_file) -> str:
    """Save observation image and attach to handover. Returns relative URL."""
    ext = Path(uploaded_file.name).suffix.lower()
    safe_name = f"{handover_id}_{uuid.uuid4().hex[:8]}{ext}"
    dest = HANDOVER_IMAGES_DIR / safe_name
    with open(dest, 'wb') as out:
        for chunk in uploaded_file.chunks():
            out.write(chunk)
    rel = f'cmms/handover_images/{safe_name}'
    # Attach to handover
    h = get_handover(handover_id)
    if h:
        imgs = h.get('observation_images', [])
        imgs.append(rel)
        update_handover(handover_id, {'observation_images': imgs})
    return rel
