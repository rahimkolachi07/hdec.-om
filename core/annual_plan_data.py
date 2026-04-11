"""
Annual Plan - equipment categories, slugs, icons, colors.
Sheet: PM Calander (gid=1182899491)
"""

ANNUAL_SHEET_ID = "1l418pkJkui6o_0Ib6r5-d0gr6AQrti_GocCr_GKxSwo"

# All sheet tabs available
ANNUAL_SHEETS = [
    {"slug": "pm-calendar",    "name": "PM Calendar",     "sheet": "PM Calander",                  "icon": "📅", "color": "#3b9eff"},
    {"slug": "overall-pm",     "name": "Overall PM",      "sheet": "Overall PM",                   "icon": "📊", "color": "#00e5c8"},
    {"slug": "pm-worksheet",   "name": "PM Worksheet",    "sheet": "PM Worksheet",                 "icon": "📋", "color": "#a259ff"},
    {"slug": "checklist",      "name": "Checklist & Procedure", "sheet": "Checklist and Procedure Status", "icon": "✅", "color": "#22c55e"},
    {"slug": "daily-plan",     "name": "Daily Plan",      "sheet": "Daily Plan",                   "icon": "📆", "color": "#f0c040"},
]

# Equipment → folder config (derived from PM Calendar sheet)
EQUIPMENT_FOLDERS = [
    {"slug": "pv-area",               "name": "PV Area",               "icon": "☀️",  "color": "#3b9eff",  "equip": "PV Area"},
    {"slug": "weather-station",       "name": "Weather Station",       "icon": "🌦️",  "color": "#00e5c8",  "equip": "Weather Station"},
    {"slug": "mvps",                  "name": "MVPS",                  "icon": "⚡",  "color": "#f0c040",  "equip": "MVPS"},
    {"slug": "tracker",               "name": "Tracker",               "icon": "📡",  "color": "#22c55e",  "equip": "Tracker"},
    {"slug": "robot",                 "name": "Robot",                 "icon": "🤖",  "color": "#a259ff",  "equip": "Robot"},
    {"slug": "module",                "name": "PV Modules",            "icon": "🔋",  "color": "#fb923c",  "equip": "Module"},
    {"slug": "scb",                   "name": "SCB",                   "icon": "🔌",  "color": "#e879f9",  "equip": "SCB"},
    {"slug": "substation",            "name": "Substation",            "icon": "🏭",  "color": "#ef4444",  "equip": "Substation"},
    {"slug": "battery-bank",          "name": "Battery Bank",          "icon": "🔋",  "color": "#4ade80",  "equip": "Battery Bank"},
    {"slug": "statcom",               "name": "STATCOM",               "icon": "⚙️",  "color": "#818cf8",  "equip": "STATCOM"},
    {"slug": "power-transformer",     "name": "Power Transformer",     "icon": "🔆",  "color": "#fbbf24",  "equip": "Power Transformer"},
    {"slug": "earthing-transformer",  "name": "Earthing Transformer",  "icon": "⬇️",  "color": "#a78bfa",  "equip": "Earthing Transformer"},
    {"slug": "gis-110kv",             "name": "110KV GIS",             "icon": "🔴",  "color": "#f87171",  "equip": "110 KV GIS"},
    {"slug": "gis-33kv",              "name": "33KV GIS",              "icon": "🟠",  "color": "#fb923c",  "equip": "33KV GIS"},
    {"slug": "ac-dc-system",          "name": "AC/DC System",          "icon": "💡",  "color": "#38bdf8",  "equip": "AC/DC System"},
    {"slug": "protection-panel",      "name": "Protection Panel",      "icon": "🛡️",  "color": "#c084fc",  "equip": "Protection Pannel"},
    {"slug": "distribution-panel",    "name": "Distribution Panel",    "icon": "📊",  "color": "#2dd4bf",  "equip": "Distribution Pannel"},
    {"slug": "hvac",                  "name": "HVAC System",           "icon": "❄️",  "color": "#7dd3fc",  "equip": "HVAC System (All Buildings)"},
    {"slug": "cctv",                  "name": "CCTV",                  "icon": "📹",  "color": "#94a3b8",  "equip": "CCTV"},
    {"slug": "main-gate",             "name": "Main Gate",             "icon": "🚪",  "color": "#d97706",  "equip": "Main Gate"},
    {"slug": "fire-extinguisher",     "name": "Fire Extinguisher",     "icon": "🧯",  "color": "#ef4444",  "equip": "Fire Extinguisher"},
    {"slug": "fire-alarm",            "name": "Fire Alarm",            "icon": "🚨",  "color": "#f97316",  "equip": "Fire Alaram"},
    {"slug": "edg",                   "name": "EDG",                   "icon": "⚙️",  "color": "#84cc16",  "equip": "EDG"},
    {"slug": "tools",                 "name": "Tools & Equipment",     "icon": "🔧",  "color": "#6b7280",  "equip": "Tools"},
    {"slug": "security",              "name": "Security",              "icon": "🔒",  "color": "#6366f1",  "equip": "Security"},
    {"slug": "general",               "name": "General",               "icon": "📋",  "color": "#64748b",  "equip": "General"},
    {"slug": "pvfd",                  "name": "PVFD",                  "icon": "💡",  "color": "#8b5cf6",  "equip": "PVFD"},
    {"slug": "pvss",                  "name": "PVSS",                  "icon": "🏢",  "color": "#0ea5e9",  "equip": "PVSS"},
]

EQUIP_SLUG_MAP  = {f["slug"]: f for f in EQUIPMENT_FOLDERS}
SHEET_SLUG_MAP  = {s["slug"]: s for s in ANNUAL_SHEETS}

FREQ_COLORS = {
    "Daily":       "#3b9eff",
    "Weekly":      "#00e5c8",
    "Monthly":     "#a259ff",
    "Quarterly":   "#f0c040",
    "Half Yearly": "#22c55e",
    "Yearly":      "#fb923c",
    "2 Yearly":    "#ef4444",
}
