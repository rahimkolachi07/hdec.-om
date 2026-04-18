from django.urls import path
from . import views
from . import project_views
from . import hse_views
from . import cmms_views
from . import meeting_views

urlpatterns = [
    path('', views.home, name='home'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('admin-panel/', views.admin_panel, name='admin_panel'),
    path('api/admin/', views.admin_api, name='admin_api'),

    # Manpower
    path('manpower/', views.manpower, name='manpower'),
    path('store/', views.legacy_store_redirect, name='legacy_store_redirect'),
    path('api/manpower/import/', views.manpower_import, name='manpower_import'),
    path('api/manpower/export/', views.manpower_export, name='manpower_export'),

    # Attendance
    path('api/attendance/face/', views.attendance_face_api, name='attendance_face_api'),
    path('api/attendance/face/delete/', views.attendance_face_delete, name='attendance_face_delete'),
    path('api/attendance/face/descriptors/', views.attendance_face_descriptors, name='attendance_face_descriptors'),
    path('api/attendance/people/', views.attendance_people, name='attendance_people'),
    path('api/attendance/face/photo/save/', views.attendance_face_photo_save, name='attendance_face_photo_save'),
    path('api/attendance/face/photo/all/', views.attendance_face_photos_all, name='attendance_face_photos_all'),
    path('api/attendance/face/photo/<str:name>/', views.attendance_face_photo_get, name='attendance_face_photo_get'),
    path('api/attendance/mark/', views.attendance_mark, name='attendance_mark'),
    path('api/attendance/', views.attendance_get, name='attendance_get'),
    path('api/attendance/export/', views.attendance_export, name='attendance_export'),

    # Tracing
    path('tracing/', views.tracing_hub, name='tracing_hub'),
    path('tracing/<slug:slug>/', views.tracing_sheet, name='tracing_sheet'),

    # Annual Plan
    path('annual-plan/', views.annual_plan, name='annual_plan'),
    path('annual-plan/sheet/<slug:slug>/', views.annual_plan_sheet, name='annual_plan_sheet'),
    path('annual-plan/<slug:slug>/', views.annual_plan_folder, name='annual_plan_folder'),

    # Other pages
    path('documents/', views.documents, name='documents'),
    path('daily-report/', views.daily_report, name='daily_report'),

    # APIs
    path('api/chat/', views.chat_api, name='chat_api'),
    path('api/tracing/gids/', views.tracing_gids_api, name='tracing_gids_api'),
    path('api/tracing/<slug:slug>/', views.tracing_sheet_api, name='tracing_sheet_api'),
    path('api/annual-plan/', views.annual_plan_api, name='annual_plan_api'),
    path('api/annual-plan/sheet/<slug:slug>/', views.annual_plan_sheet_api, name='annual_plan_sheet_api'),
    path('api/annual-plan/<slug:slug>/', views.annual_plan_folder_api, name='annual_plan_folder_api'),

    # ── Multi-Country / Multi-Project ────────────────────────────────────
    path('c/<str:country_id>/', views.country_view, name='country_view'),
    path('p/<str:country_id>/<str:project_id>/', views.project_hub_view, name='project_hub'),
    path('api/projects/', views.projects_api, name='projects_api'),
    path('api/projects/reorder/', views.projects_reorder_api, name='projects_reorder_api'),

    # ── Meeting Room ─────────────────────────────────────────────────────
    # Keep these before the generic category route so "meeting-room"
    # is treated as the dedicated module path, not as a category slug.
    path('p/<str:country_id>/<str:project_id>/meeting-room/', meeting_views.meeting_hub, name='meeting_hub'),
    path('meeting-room/', meeting_views.meeting_hub, name='meeting_hub_direct'),
    path('api/meeting/poll/',                     meeting_views.meeting_api_poll,          name='meeting_api_poll'),
    path('api/meeting/presence/',                 meeting_views.meeting_api_presence,      name='meeting_api_presence'),
    path('api/meeting/users/',                    meeting_views.meeting_api_users,         name='meeting_api_users'),
    path('api/meeting/rooms/',                    meeting_views.meeting_api_rooms,         name='meeting_api_rooms'),
    path('api/meeting/rooms/<str:room_id>/',      meeting_views.meeting_api_room,          name='meeting_api_room'),
    path('api/meeting/rooms/<str:room_id>/join/', meeting_views.meeting_api_room_join,     name='meeting_api_room_join'),
    path('api/meeting/rooms/<str:room_id>/leave/',meeting_views.meeting_api_room_leave,    name='meeting_api_room_leave'),
    path('api/meeting/messages/',                 meeting_views.meeting_api_messages,      name='meeting_api_messages'),
    path('api/meeting/groups/',                   meeting_views.meeting_api_groups,        name='meeting_api_groups'),
    path('api/meeting/groups/<str:group_id>/',    meeting_views.meeting_api_group,         name='meeting_api_group'),
    path('api/meeting/files/',                    meeting_views.meeting_api_files,         name='meeting_api_files'),
    path('api/meeting/files/<str:file_id>/',      meeting_views.meeting_api_file_download, name='meeting_api_file_download'),
    path('api/meeting/calls/',                    meeting_views.meeting_api_calls,         name='meeting_api_calls'),
    path('api/meeting/realtime/session/',         meeting_views.meeting_api_realtime_session, name='meeting_api_realtime_session'),
    path('api/meeting/calls/<str:call_id>/',      meeting_views.meeting_api_call,          name='meeting_api_call'),
    path('api/meeting/room-meetings/',            meeting_views.meeting_api_room_meetings, name='meeting_api_room_meetings'),
    path('api/meeting/room-meetings/<str:meeting_id>/', meeting_views.meeting_api_room_meeting, name='meeting_api_room_meeting'),
    path('api/meeting/room-signals/',             meeting_views.meeting_api_room_signals,  name='meeting_api_room_signals'),
    path('api/meeting/global-alerts/',            meeting_views.meeting_api_global_alerts, name='meeting_api_global_alerts'),

    # ── Category Hub (Country → Project → Category → Modules) ────────────
    path('p/<str:country_id>/<str:project_id>/<str:category>/', views.category_hub_view, name='category_hub'),

    # ── Category-scoped: Manpower ─────────────────────────────────────────
    path('p/<str:country_id>/<str:project_id>/<str:category>/manpower/', project_views.project_manpower, name='project_manpower'),
    path('api/p/<str:country_id>/<str:project_id>/<str:category>/manpower/', project_views.project_manpower_api, name='project_manpower_api'),

    # ── Category-scoped: Store ────────────────────────────────────────────
    path('p/<str:country_id>/<str:project_id>/<str:category>/store/', project_views.project_store, name='project_store'),
    path('api/p/<str:country_id>/<str:project_id>/<str:category>/store/', project_views.project_store_api, name='project_store_api'),

    # ── CMMS ─────────────────────────────────────────────────────────────
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/', project_views.project_cmms_hub, name='project_cmms_hub'),
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/activities/', project_views.project_cmms_activities, name='project_cmms_activities'),
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/permits/', project_views.project_cmms_permits, name='project_cmms_permits'),
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/permits/new/', project_views.project_cmms_permit_new, name='project_cmms_permit_new'),
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/permits/<str:permit_id>/', project_views.project_cmms_permit_detail, name='project_cmms_permit_detail'),
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/handover/', project_views.project_cmms_handover_list, name='project_cmms_handover_list'),
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/handover/new/', project_views.project_cmms_handover_new, name='project_cmms_handover_new'),
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/handover/<str:handover_id>/', project_views.project_cmms_handover_detail, name='project_cmms_handover_detail'),
    path('api/p/<str:country_id>/<str:project_id>/<str:category>/handover/', project_views.project_cmms_handover_api, name='project_cmms_handover_api'),
    path('cmms/', cmms_views.cmms_hub, name='cmms_hub'),
    path('cmms/activities/', cmms_views.cmms_hub, name='cmms_activities_legacy'),
    path('cmms/permits/', cmms_views.cmms_ptw_list, name='cmms_permits_legacy'),
    path('cmms/permits/<str:permit_id>/', cmms_views.cmms_ptw_detail, name='cmms_permit_detail_legacy'),
    path('cmms/permits/<str:permit_id>/download/', cmms_views.cmms_ptw_download, name='cmms_permit_download_legacy'),
    path('cmms/handover/', cmms_views.cmms_handover_legacy, name='cmms_handover_legacy'),
    path('cmms/ptw/', cmms_views.cmms_ptw_list, name='cmms_ptw_list'),
    path('cmms/ptw/<str:permit_id>/', cmms_views.cmms_ptw_detail, name='cmms_ptw_detail'),
    path('cmms/ptw/<str:permit_id>/download/', cmms_views.cmms_ptw_download, name='cmms_ptw_download'),
    path('cmms/work/<str:record_id>/', cmms_views.cmms_work, name='cmms_work'),
    path('cmms/zip/<str:record_id>/', cmms_views.cmms_download_zip, name='cmms_zip'),

    # CMMS APIs
    path('api/cmms/activities/', cmms_views.cmms_api_activities, name='cmms_api_activities'),
    path('api/cmms/activities/<str:activity_id>/', cmms_views.cmms_api_activity, name='cmms_api_activity'),
    path('api/cmms/start/', cmms_views.cmms_api_start, name='cmms_api_start'),
    path('api/cmms/ptw/<str:permit_id>/', cmms_views.cmms_api_ptw, name='cmms_api_ptw'),
    path('api/notifications/', cmms_views.notifications_api, name='notifications_api'),
    path('api/cmms/excel/<str:record_id>/', cmms_views.cmms_api_excel, name='cmms_api_excel'),
    path('api/cmms/photos/<str:record_id>/', cmms_views.cmms_api_photos, name='cmms_api_photos'),
    path('api/cmms/complete/<str:record_id>/', cmms_views.cmms_api_complete, name='cmms_api_complete'),
    path('api/cmms/checklists/', cmms_views.cmms_api_checklists, name='cmms_api_checklists'),
    path('api/cmms/checklist-activities/', cmms_views.cmms_api_checklist_activities, name='cmms_api_checklist_activities'),
    path('cmms/checklist/<str:record_id>/', cmms_views.cmms_checklist_native, name='cmms_checklist_native'),
    path('cmms/report/<str:record_id>/', cmms_views.cmms_report_native, name='cmms_report_native'),
    path('api/cmms/checklist/<str:record_id>/', cmms_views.cmms_api_checklist_save, name='cmms_api_checklist_save'),
    path('api/cmms/report/<str:record_id>/', cmms_views.cmms_api_report_save, name='cmms_api_report_save'),

    # ── HSE ──────────────────────────────────────────────────────────────
    path('hse/sjn-portal/', views.hse_sjn_portal, name='hse_sjn_portal'),
    path('api/hse/permits/', hse_views.hse_api_permits, name='hse_api_permits'),
    path('api/hse/permits/<str:permit_id>/', hse_views.hse_api_permit_detail, name='hse_api_permit_detail'),
    path('api/hse/records/', hse_views.hse_api_records, name='hse_api_records'),
    path('api/hse/records/<str:record_id>/', hse_views.hse_api_record_detail, name='hse_api_record_detail'),

    # ── Admin Management Modules ──────────────────────────────────────────────
    path('api/admin/vehicles/', views.admin_vehicles_api, name='admin_vehicles'),
    path('api/admin/vehicles/<str:rid>/', views.admin_vehicle_api, name='admin_vehicle'),
    path('api/admin/residences/', views.admin_residences_api, name='admin_residences'),
    path('api/admin/residences/<str:rid>/', views.admin_residence_api, name='admin_residence'),
    path('api/admin/workforce/', views.admin_workforce_api, name='admin_workforce'),
    path('api/admin/workforce/<str:rid>/', views.admin_workforce_member_api, name='admin_workforce_member'),
    path('api/admin/gate-passes/', views.admin_gatepasses_api, name='admin_gatepasses'),
    path('api/admin/gate-passes/<str:rid>/', views.admin_gatepass_api, name='admin_gatepass'),
    path('api/admin/equipment/', views.admin_equipment_api, name='admin_equipment'),
    path('api/admin/equipment/<str:rid>/', views.admin_equipment_item_api, name='admin_equipment_item'),
    path('api/admin/trainings/', views.admin_trainings_api, name='admin_trainings'),
    path('api/admin/trainings/<str:rid>/', views.admin_training_api, name='admin_training'),
]
