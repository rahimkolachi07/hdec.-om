from django.urls import path
from . import views
from . import cmms_views
from . import project_views

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
    path('api/tracing/<slug:slug>/', views.tracing_sheet_api, name='tracing_sheet_api'),
    path('api/annual-plan/', views.annual_plan_api, name='annual_plan_api'),
    path('api/annual-plan/sheet/<slug:slug>/', views.annual_plan_sheet_api, name='annual_plan_sheet_api'),
    path('api/annual-plan/<slug:slug>/', views.annual_plan_folder_api, name='annual_plan_folder_api'),

    # ── Multi-Country / Multi-Project ────────────────────────────────────
    path('c/<str:country_id>/', views.country_view, name='country_view'),
    path('p/<str:country_id>/<str:project_id>/', views.project_hub_view, name='project_hub'),
    path('api/projects/', views.projects_api, name='projects_api'),

    # ── Category Hub (Country → Project → Category → Modules) ────────────
    path('p/<str:country_id>/<str:project_id>/<str:category>/', views.category_hub_view, name='category_hub'),

    # ── Category-scoped: Manpower ─────────────────────────────────────────
    path('p/<str:country_id>/<str:project_id>/<str:category>/manpower/', project_views.project_manpower, name='project_manpower'),
    path('api/p/<str:country_id>/<str:project_id>/<str:category>/manpower/', project_views.project_manpower_api, name='project_manpower_api'),

    # ── Category-scoped: Store ────────────────────────────────────────────
    path('p/<str:country_id>/<str:project_id>/<str:category>/store/', project_views.project_store, name='project_store'),
    path('api/p/<str:country_id>/<str:project_id>/<str:category>/store/', project_views.project_store_api, name='project_store_api'),

    # ── Category-scoped: CMMS Hub ─────────────────────────────────────────
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/', project_views.project_cmms_hub, name='project_cmms_hub'),

    # ── Category-scoped: Activities ───────────────────────────────────────
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/activities/', project_views.project_activities, name='project_activities'),
    path('api/p/<str:country_id>/<str:project_id>/<str:category>/activities/', project_views.project_activities_api, name='project_activities_api'),

    # ── Category-scoped: Permits ──────────────────────────────────────────
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/permits/', project_views.project_permits, name='project_permits'),
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/permits/new/', project_views.project_permit_detail, name='project_permit_new'),
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/permits/<str:permit_id>/', project_views.project_permit_detail, name='project_permit_detail'),
    path('api/p/<str:country_id>/<str:project_id>/<str:category>/permits/', project_views.project_permits_api, name='project_permits_api'),

    # ── Category-scoped: Handover ─────────────────────────────────────────
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/handover/', project_views.project_handovers, name='project_handovers'),
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/handover/new/', project_views.project_handover_detail, name='project_handover_new'),
    path('p/<str:country_id>/<str:project_id>/<str:category>/cmms/handover/<str:handover_id>/', project_views.project_handover_detail, name='project_handover_detail'),
    path('api/p/<str:country_id>/<str:project_id>/<str:category>/handover/', project_views.project_handover_api, name='project_handover_api'),

    # ── CMMS ─────────────────────────────────────────────────────────────
    path('cmms/', cmms_views.cmms_hub, name='cmms_hub'),

    # Activities
    path('cmms/activities/', cmms_views.cmms_activities, name='cmms_activities'),
    path('cmms/activities/<str:activity_id>/', cmms_views.cmms_activity_detail, name='cmms_activity_detail'),
    path('cmms/activities/<str:activity_id>/zip/', cmms_views.cmms_download_activity_zip, name='cmms_activity_zip'),

    # Records (per activity per day)
    path('cmms/records/<str:record_id>/zip/', cmms_views.cmms_download_zip, name='cmms_record_zip'),
    path('api/cmms/checklist/<str:record_id>/', cmms_views.cmms_checklist_api, name='cmms_checklist_api'),
    path('api/cmms/photos/<str:record_id>/', cmms_views.cmms_photo_api, name='cmms_photo_api'),

    # Activities admin
    path('api/cmms/activities/', cmms_views.cmms_activity_api, name='cmms_activity_api'),

    # Permits
    path('cmms/permits/', cmms_views.cmms_permits, name='cmms_permits'),
    path('cmms/permits/new/', cmms_views.cmms_permit_detail, name='cmms_permit_new'),
    path('cmms/permits/<str:permit_id>/', cmms_views.cmms_permit_detail, name='cmms_permit_detail'),
    path('cmms/permits/<str:permit_id>/pdf/', cmms_views.cmms_permit_pdf, name='cmms_permit_pdf'),

    # Permit API
    path('api/cmms/permits/', cmms_views.cmms_permit_api, name='cmms_permit_api'),

    # Email config (admin)
    path('api/cmms/email-config/', cmms_views.cmms_email_config_api, name='cmms_email_config_api'),

    # Manpower duty staff API
    path('api/cmms/duty-staff/', cmms_views.cmms_duty_staff_api, name='cmms_duty_staff_api'),

    # ICC PDF download
    path('cmms/permits/<str:permit_id>/icc/', cmms_views.cmms_icc_pdf, name='cmms_icc_pdf'),

    # Word (.docx) permit download
    path('cmms/permits/<str:permit_id>/docx/', cmms_views.cmms_permit_docx, name='cmms_permit_docx'),

    # Activity email trigger
    path('api/cmms/activity-email/<str:record_id>/', cmms_views.cmms_send_activity_email, name='cmms_send_activity_email'),

    # ── Handover / Shift Log ──────────────────────────────────────────────
    path('cmms/handover/', cmms_views.cmms_handovers, name='cmms_handovers'),
    path('cmms/handover/new/', cmms_views.cmms_handover_detail, name='cmms_handover_new'),
    path('cmms/handover/<str:handover_id>/', cmms_views.cmms_handover_detail, name='cmms_handover_detail'),
    path('api/cmms/handover/', cmms_views.cmms_handover_api, name='cmms_handover_api'),
    path('api/cmms/handover/image/<str:handover_id>/', cmms_views.cmms_handover_image_api, name='cmms_handover_image_api'),
]
