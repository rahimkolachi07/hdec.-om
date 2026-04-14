"""
project_utils.py — Multi-Country / Multi-Project / Multi-Category configuration store.

Data lives in  <project_root>/projects.json  (no database needed).

Structure:
  {
    "countries": [
      {
        "id": "saudi_arabia",
        "name": "Saudi Arabia",
        "flag": "🇸🇦",
        "color": "#006c35",
        "projects": [
          {
            "id": "henakiya_1",
            "name": "1100MW Al-Henakiya CCPP Ph-1",
            "description": "...",
            "legacy": true,
            "categories": {
              "maintenance": { "modules": ["activities", "permits", ...] },
              "operation":   { "modules": [] },
              "construction":{ "modules": [] }
            },
            "created_at": "2024-01-01T00:00:00"
          }
        ]
      }
    ]
  }
"""
import json, re
from pathlib import Path
from datetime import datetime

PROJECTS_FILE = Path(__file__).resolve().parent.parent / 'projects.json'

# ── Category metadata ──────────────────────────────────────────────────────────

CATEGORY_META = {
    'maintenance': {
        'label': 'Maintenance',
        'icon': '🔧',
        'color': '#3b9eff',
        'color_rgb': '59,158,255',
        'desc': 'CMMS, Permits, Shift Handover, Manpower and maintenance planning',
    },
    'operation': {
        'label': 'Operation',
        'icon': '⚡',
        'color': '#00e5c8',
        'color_rgb': '0,229,200',
        'desc': 'Daily operations, tracing sheets and operational reporting',
    },
    'construction': {
        'label': 'Construction',
        'icon': '🏗️',
        'color': '#f0c040',
        'color_rgb': '240,192,64',
        'desc': 'Construction activities, documents and progress tracking',
    },
    'hse': {
        'label': 'HSE',
        'icon': '🦺',
        'color': '#22c55e',
        'color_rgb': '34,197,94',
        'desc': 'Health, Safety and Environment management, permits and KPI tracking',
    },
}

ALL_CATEGORY_IDS = ['maintenance', 'operation', 'construction', 'hse']

# ── Module metadata ────────────────────────────────────────────────────────────
MODULE_META = {
    'activities': {
        'label': 'CMMS Activities',
        'icon': '📋',
        'color': '#3b9eff',
        'desc': 'Maintenance activity records, checklists & photo logs',
        'route': '/cmms/',
        'hub_route': '/cmms/',
        'category': 'maintenance',
    },
    'permits': {
        'label': 'Work Permits (PTW)',
        'icon': '🔐',
        'color': '#a259ff',
        'desc': 'Permit to Work application, issuance & closure workflow',
        'route': '/cmms/permits/',
        'hub_route': '/cmms/',
        'category': 'maintenance',
    },
    'handover': {
        'label': 'Shift Handover',
        'icon': '📝',
        'color': '#00e5c8',
        'desc': 'Shift log entries and inter-shift communication records',
        'route': '/cmms/handover/',
        'hub_route': '/cmms/',
        'category': 'maintenance',
    },
    'manpower': {
        'label': 'Manpower',
        'icon': '👷',
        'color': '#f0c040',
        'desc': 'Workforce scheduling, duty roster and attendance tracking',
        'route': '/manpower/',
        'hub_route': '/manpower/',
        'category': 'maintenance',
    },
    'store': {
        'label': 'Store',
        'icon': '📦',
        'color': '#ef4444',
        'desc': 'Equipment issue and return tracking with photos and stock details',
        'route': None,
        'hub_route': None,
        'category': 'maintenance',
    },
    'tracing': {
        'label': 'Tracing Sheets',
        'icon': '🗂️',
        'color': '#22c55e',
        'desc': 'Equipment tracing, PM/CM statistics and inspection sheets',
        'route': '/tracing/',
        'hub_route': '/tracing/',
        'category': 'maintenance',
    },
    'annual_plan': {
        'label': 'Annual Plan',
        'icon': '📅',
        'color': '#e879f9',
        'desc': 'Yearly maintenance schedule and planning sheets',
        'route': '/annual-plan/',
        'hub_route': '/annual-plan/',
        'category': 'maintenance',
    },
    'documents': {
        'label': 'Documents',
        'icon': '📄',
        'color': '#38bdf8',
        'desc': 'Project documentation and technical reference library',
        'route': '/documents/',
        'hub_route': '/documents/',
        'category': 'maintenance',
    },
    'daily_report': {
        'label': 'Daily Report',
        'icon': '📊',
        'color': '#fb923c',
        'desc': 'Daily operations summary and performance tracking',
        'route': '/daily-report/',
        'hub_route': '/daily-report/',
        'category': 'maintenance',
    },
    'sjn_portal': {
        'label': 'SJN Portal',
        'icon': '🌿',
        'color': '#22c55e',
        'desc': 'O&M portal — Permits, DPR, Gate Pass, LOTO, SCADA, HSSE KPI and Asset Management',
        'route': '/hse/sjn-portal/',
        'hub_route': '/hse/sjn-portal/',
        'category': 'hse',
    },
}

ALL_MODULE_IDS = list(MODULE_META.keys())


# ── JSON helpers ───────────────────────────────────────────────────────────────

def _load() -> dict:
    if not PROJECTS_FILE.exists():
        data = {'countries': []}
        _save(data)
        return data
    try:
        with open(PROJECTS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {'countries': []}


def _save(data: dict):
    with open(PROJECTS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def _slugify(text: str) -> str:
    """Convert text to a safe slug id."""
    s = text.lower().strip()
    s = re.sub(r'[^\w\s-]', '', s)
    s = re.sub(r'[\s-]+', '_', s)
    return s.strip('_')


def _unique_id(base: str, existing: list) -> str:
    candidate = _slugify(base)
    if candidate not in existing:
        return candidate
    counter = 2
    while f'{candidate}_{counter}' in existing:
        counter += 1
    return f'{candidate}_{counter}'


# ── Category helpers ───────────────────────────────────────────────────────────

def _migrate_project_categories(project: dict) -> dict:
    """
    Auto-migrate a project that uses the old flat `modules` list to the new
    categories structure. Returns the categories dict (does NOT mutate project).
    """
    if 'categories' in project:
        return project['categories']
    # Old format: all modules belong to maintenance
    return {
        'maintenance': {'modules': project.get('modules', [])},
        'operation':   {'modules': []},
        'construction': {'modules': []},
        'hse':         {'modules': []},
    }


def get_project_categories(project: dict) -> dict:
    """Return categories dict, auto-migrating from old flat modules format."""
    cats = _migrate_project_categories(project)
    normalized = {}
    for cid in ALL_CATEGORY_IDS:
        src = cats.get(cid, {})
        modules = [m for m in src.get('modules', []) if m in MODULE_META]
        if cid == 'maintenance' and 'store' not in modules:
            modules.append('store')
        normalized[cid] = {'modules': modules}
    return normalized


def get_category_modules(project: dict, category: str) -> list:
    """Return module ids for a specific category in a project."""
    cats = get_project_categories(project)
    return cats.get(category, {}).get('modules', [])


def get_all_modules_flat(project: dict) -> list:
    """Return all module ids across all categories (for counts/display)."""
    cats = get_project_categories(project)
    result = []
    for cat_data in cats.values():
        result.extend(cat_data.get('modules', []))
    return result


def _blank_categories(category: str = None, modules: list = None) -> dict:
    """Return a fresh categories dict, optionally seeding one category with modules."""
    cats = {cid: {'modules': []} for cid in ALL_CATEGORY_IDS}
    if category and modules is not None:
        cats[category]['modules'] = [m for m in modules if m in MODULE_META]
    return cats


# ── Country CRUD ───────────────────────────────────────────────────────────────

def get_countries() -> list:
    return _load().get('countries', [])


def get_country(country_id: str) -> dict | None:
    for c in get_countries():
        if c['id'] == country_id:
            return c
    return None


def create_country(name: str, flag: str = '🏴', color: str = '#3b9eff') -> str:
    """Create a new country; returns the new country id."""
    data = _load()
    existing = [c['id'] for c in data['countries']]
    new_id = _unique_id(name, existing)
    data['countries'].append({
        'id': new_id,
        'name': name.strip(),
        'flag': flag.strip(),
        'color': color.strip(),
        'projects': [],
    })
    _save(data)
    return new_id


def update_country(country_id: str, fields: dict):
    data = _load()
    for c in data['countries']:
        if c['id'] == country_id:
            for k, v in fields.items():
                if k not in ('id', 'projects'):
                    c[k] = v
    _save(data)


def delete_country(country_id: str):
    data = _load()
    data['countries'] = [c for c in data['countries'] if c['id'] != country_id]
    _save(data)


# ── Project CRUD ───────────────────────────────────────────────────────────────

def get_project(country_id: str, project_id: str) -> dict | None:
    c = get_country(country_id)
    if not c:
        return None
    for p in c.get('projects', []):
        if p['id'] == project_id:
            return p
    return None


def create_project(country_id: str, name: str, description: str = '',
                   categories: dict = None, legacy: bool = False) -> str | None:
    """
    Create a project under a country.

    `categories` can be:
      - None  → all modules in maintenance (default)
      - dict  → {'maintenance': {'modules': [...]}, 'operation': {...}, ...}

    Returns the new project id or None if country not found.
    """
    data = _load()
    for c in data['countries']:
        if c['id'] != country_id:
            continue
        existing = [p['id'] for p in c.get('projects', [])]
        new_id = _unique_id(name, existing)
        if categories is None:
            cats = _blank_categories('maintenance', ALL_MODULE_IDS[:])
        else:
            # Validate module ids in each category
            cats = {}
            for cid in ALL_CATEGORY_IDS:
                cat_data = categories.get(cid, {})
                cats[cid] = {
                    'modules': [m for m in cat_data.get('modules', []) if m in MODULE_META]
                }
        c.setdefault('projects', []).append({
            'id': new_id,
            'name': name.strip(),
            'description': description.strip(),
            'legacy': legacy,
            'categories': cats,
            'created_at': datetime.now().isoformat(),
        })
        _save(data)
        return new_id
    return None


def update_project(country_id: str, project_id: str, fields: dict):
    data = _load()
    for c in data['countries']:
        if c['id'] != country_id:
            continue
        for p in c.get('projects', []):
            if p['id'] == project_id:
                for k, v in fields.items():
                    if k != 'id':
                        p[k] = v
    _save(data)


def delete_project(country_id: str, project_id: str):
    data = _load()
    for c in data['countries']:
        if c['id'] == country_id:
            c['projects'] = [p for p in c.get('projects', []) if p['id'] != project_id]
    _save(data)


# ── Module cards helper ────────────────────────────────────────────────────────

def get_project_module_cards(project: dict, user_permissions: dict = None,
                             country_id: str = '', project_id: str = '',
                             category: str = 'maintenance') -> list:
    """
    Return a list of module card dicts for a specific category within the project,
    filtered by user permissions (modules with 'none' access are excluded).
    """
    # Project-scoped route map: category-prefixed URLs
    def _route(mod_id):
        routes = {
            'manpower':    f'/p/{country_id}/{project_id}/{category}/manpower/',
            'store':       f'/p/{country_id}/{project_id}/{category}/store/',
            'activities':  f'/p/{country_id}/{project_id}/{category}/cmms/activities/',
            'permits':     f'/p/{country_id}/{project_id}/{category}/cmms/permits/',
            'handover':    f'/p/{country_id}/{project_id}/{category}/cmms/handover/',
            'tracing':     None,
            'annual_plan': None,
            'documents':   None,
            'daily_report': None,
            'sjn_portal':  '/hse/sjn-portal/',
        }
        return routes.get(mod_id)

    module_ids = get_category_modules(project, category)
    cards = []
    for mod_id in module_ids:
        meta = MODULE_META.get(mod_id)
        if not meta:
            continue
        access = 'view'
        if user_permissions:
            access = user_permissions.get(mod_id, 'none')
        if access == 'none':
            continue
        if project.get('legacy') and meta.get('route'):
            route = meta['route']
        else:
            route = _route(mod_id)
        cards.append({
            'id': mod_id,
            'label': meta['label'],
            'icon': meta['icon'],
            'color': meta['color'],
            'desc': meta['desc'],
            'route': route,
            'can_edit': access == 'edit',
            'legacy': project.get('legacy', False),
        })
    return cards
