"""
Email notification utilities for HDEC CMMS.
Sends notifications for the permit workflow, activity assignments,
and the daily activity digest.
"""
import json
import logging
from datetime import datetime, timedelta
from html import escape
from pathlib import Path

from django.conf import settings
from django.core.mail import send_mail

logger = logging.getLogger(__name__)

EMAIL_AUDIT_FILE = Path(getattr(settings, 'CMMS_DATA_DIR', Path.cwd() / 'cmms_data')) / 'email_audit.jsonl'
PORTAL_BASE_URL = str(getattr(settings, 'HDEC_PORTAL_URL', 'https://hdec-om.live')).rstrip('/')


def _write_email_audit(entry: dict):
    try:
        EMAIL_AUDIT_FILE.parent.mkdir(parents=True, exist_ok=True)
        with EMAIL_AUDIT_FILE.open('a', encoding='utf-8') as handle:
            handle.write(json.dumps(entry, ensure_ascii=True) + '\n')
    except Exception as exc:
        logger.warning(f"[CMMS Email] Failed to write audit log: {exc}")


def _send(to_emails: list, subject: str, text: str, html: str = None):
    """Send email to a list of recipients and write an audit entry."""
    recipients = list(dict.fromkeys(
        str(email).strip()
        for email in to_emails
        if email and '@' in str(email)
    ))
    if not recipients:
        logger.info(f"[CMMS Email] No valid recipients for: {subject}")
        _write_email_audit({
            'timestamp': datetime.now().isoformat(),
            'subject': subject,
            'recipients': [],
            'status': 'skipped_no_recipients',
        })
        return 0

    try:
        sent_count = send_mail(
            subject=subject,
            message=text,
            from_email=getattr(settings, 'DEFAULT_FROM_EMAIL', 'HDEC CMMS <noreply@hdec.sa>'),
            recipient_list=recipients,
            html_message=html,
            fail_silently=False,
        )
        logger.info(f"[CMMS Email] Sent '{subject}' to {recipients}")
        _write_email_audit({
            'timestamp': datetime.now().isoformat(),
            'subject': subject,
            'recipients': recipients,
            'status': 'sent',
            'sent_count': sent_count,
        })
        return sent_count
    except Exception as exc:
        logger.warning(f"[CMMS Email] Failed to send '{subject}': {exc}")
        _write_email_audit({
            'timestamp': datetime.now().isoformat(),
            'subject': subject,
            'recipients': recipients,
            'status': 'failed',
            'error': str(exc),
        })
        return 0


def _portal_home_url() -> str:
    return f"{PORTAL_BASE_URL}/"


def _permit_portal_url(permit: dict) -> str:
    permit_id = str((permit or {}).get('id', '')).strip()
    if permit_id:
        return f"{PORTAL_BASE_URL}/cmms/ptw/{permit_id}/"
    return _portal_home_url()


def _safe(value, default: str = '-') -> str:
    text = str(value or '').strip()
    return text or default


def _activity_details(permit: dict) -> dict:
    details = {
        'name': permit.get('activity_name', ''),
        'description': permit.get('activity_description', ''),
        'frequency': permit.get('activity_frequency', ''),
        'scheduled_date': permit.get('scheduled_date', ''),
        'assigned_engineer': permit.get('assigned_engineer', ''),
        'assigned_technician': permit.get('assigned_technician', ''),
    }

    activity_id = str((permit or {}).get('activity_id', '')).strip()
    if activity_id:
        try:
            from .cmms_utils import get_activity
            activity = get_activity(activity_id) or {}
        except Exception:
            activity = {}

        details['name'] = details['name'] or activity.get('name', '')
        details['description'] = details['description'] or activity.get('notes', '')
        details['frequency'] = details['frequency'] or activity.get('frequency', '')
        details['scheduled_date'] = details['scheduled_date'] or activity.get('scheduled_date', '')
        details['assigned_engineer'] = details['assigned_engineer'] or activity.get('assigned_engineer', '')
        details['assigned_technician'] = details['assigned_technician'] or activity.get('assigned_technician', '')

    return details


def _permit_subject_target(permit: dict) -> str:
    return _safe(permit.get('activity_name') or permit.get('equipment'), 'Permit')


def _final_pdf_html(permit: dict) -> str:
    link = str((permit or {}).get('final_pdf_url', '')).strip()
    if not link:
        return ''
    escaped_link = escape(link)
    return (
        f'<p style="font-size:13px;margin:16px 0 0">'
        f'<strong>Final Locked PDF:</strong> '
        f'<a href="{escaped_link}" target="_blank" rel="noopener" '
        f'style="color:#0f4aa3;text-decoration:none">{escaped_link}</a></p>'
    )


def _action_panel(permit: dict, action_label: str) -> str:
    permit_url = escape(_permit_portal_url(permit))
    portal_url = escape(_portal_home_url())
    return f"""
<div style="margin:20px 0;padding:18px;border:1px solid #dbe4f0;border-radius:14px;background:#f8fbff">
  <div style="font-size:15px;font-weight:700;color:#0f172a;margin-bottom:6px">Approval Website</div>
  <div style="font-size:13px;color:#475569;margin-bottom:14px">Use the HDEC CMMS portal below to review this permit and continue the approval process.</div>
  <a href="{permit_url}" style="display:inline-block;background:#0f4aa3;color:#ffffff;text-decoration:none;padding:12px 18px;border-radius:10px;font-size:13px;font-weight:700">{escape(action_label)}</a>
  <div style="margin-top:12px;font-size:12px;color:#64748b">
    Portal: <a href="{portal_url}" style="color:#0f4aa3;text-decoration:none">{portal_url}</a>
  </div>
</div>"""


def _html_wrap(title: str, body_html: str, color: str = '#1a3a6b') -> str:
    portal_url = escape(_portal_home_url())
    return f"""
<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background:#f4f6fb;font-family:Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fb;padding:30px 0">
  <tr>
    <td align="center">
      <table width="640" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 8px 30px rgba(15,23,42,.10)">
        <tr>
          <td style="background:{color};padding:24px 32px">
            <p style="margin:0;color:#dbeafe;font-size:11px;letter-spacing:1px;text-transform:uppercase">HDEC - 1100MW Al Henakiya Project</p>
            <h2 style="margin:6px 0 0;color:#ffffff;font-size:20px">{escape(title)}</h2>
          </td>
        </tr>
        <tr>
          <td style="padding:30px 32px;color:#2d3748;font-size:14px;line-height:1.7">
            {body_html}
            <hr style="border:none;border-top:1px solid #e2e8f0;margin:24px 0">
            <p style="font-size:12px;color:#64748b;margin:0 0 6px">Please log in to <strong>HDEC CMMS</strong> to take action.</p>
            <p style="font-size:12px;color:#64748b;margin:0">Portal: <a href="{portal_url}" style="color:#0f4aa3;text-decoration:none">{portal_url}</a></p>
          </td>
        </tr>
        <tr>
          <td style="background:#f8fafc;padding:14px 32px;font-size:11px;color:#94a3b8;text-align:center">
            HDEC CMMS - Automated Notification - Do not reply to this email
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>"""


def _permit_info_table(permit: dict) -> str:
    """Reusable HTML table block with permit and activity details."""
    activity = _activity_details(permit)
    rows = [
        ('Activity Name', _safe(activity.get('name'), 'N/A')),
        ('Activity Description', _safe(activity.get('description'), 'N/A')),
        ('Permit Receiver', _safe(permit.get('receiver_name'), 'N/A')),
        ('Equipment', _safe(permit.get('equipment'), 'N/A')),
        ('Location', _safe(permit.get('location'), 'N/A')),
        ('Work Type', _safe((permit.get('work_type') or '').replace('_', ' ').title(), 'N/A')),
        ('Scheduled Date', _safe(activity.get('scheduled_date'), 'N/A')),
        ('Frequency', _safe(activity.get('frequency'), 'N/A')),
        ('Assigned Engineer', _safe(activity.get('assigned_engineer'), 'N/A')),
        ('Assigned Technician', _safe(activity.get('assigned_technician'), 'N/A')),
        ('Applied On', _safe((permit.get('created_at', '') or '')[:10], 'N/A')),
    ]

    html_rows = []
    for index, (label, value) in enumerate(rows):
        background = '#f8fbff' if index % 2 == 0 else '#ffffff'
        html_rows.append(
            f'<tr style="background:{background}">'
            f'<td style="padding:10px 14px;font-weight:600;width:34%;color:#334155">{escape(label)}</td>'
            f'<td style="padding:10px 14px;color:#0f172a">{escape(value)}</td>'
            f'</tr>'
        )
    return f'<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:13px">{"".join(html_rows)}</table>'


def notify_permit_created(permit: dict, operation_engineers: list):
    """Step 1: Notify operation engineers of a new permit application."""
    activity = _activity_details(permit)
    applied_date = (permit.get('created_at', '') or '')[:10]
    permit_url = _permit_portal_url(permit)
    emails = [u.get('email', '') for u in operation_engineers]
    subject = f"[PTW] Action Required - New Permit: {_permit_subject_target(permit)} | {applied_date}"
    text = (
        "New Permit Application\n\n"
        f"Activity: {_safe(activity.get('name'), 'N/A')}\n"
        f"Description: {_safe(activity.get('description'), 'N/A')}\n"
        f"Receiver: {_safe(permit.get('receiver_name'), 'N/A')}\n"
        f"Equipment: {_safe(permit.get('equipment'), 'N/A')}\n"
        f"Location: {_safe(permit.get('location'), 'N/A')}\n"
        f"Approval Website: {permit_url}\n"
        "\n"
        "Please log in and issue the permit."
    )
    html_body = f"""
<p style="margin:0 0 10px">A new <strong>Permit to Work</strong> has been submitted and is awaiting your issuance.</p>
{_permit_info_table(permit)}
{_action_panel(permit, 'Open PTW For Issuer Review')}
<p style="background:#fff7ed;border-left:4px solid #f97316;padding:12px 16px;border-radius:6px;font-size:13px;margin-top:16px">
  <strong>Action Required - Permit Issuer:</strong> Review the permit in HDEC CMMS and proceed it to HSE.
</p>"""
    _send(emails, subject, text, _html_wrap('New Permit Application - Awaiting Issuance', html_body, '#1e40af'))


def notify_permit_issued(permit: dict, hse_officers: list):
    """Step 2: Notify HSE officers that the permit requires HSE approval."""
    activity = _activity_details(permit)
    permit_url = _permit_portal_url(permit)
    emails = [u.get('email', '') for u in hse_officers]
    subject = f"[HDEC CMMS] HSE Sign-off Required - Permit: {_permit_subject_target(permit)}"
    text = (
        "Work Permit Issued - Awaiting HSE Sign-off\n\n"
        f"Activity: {_safe(activity.get('name'), 'N/A')}\n"
        f"Description: {_safe(activity.get('description'), 'N/A')}\n"
        f"Receiver: {_safe(permit.get('receiver_name'), 'N/A')}\n"
        f"Issuer: {_safe(permit.get('issuer_name'), 'N/A')}\n"
        f"Equipment: {_safe(permit.get('equipment'), 'N/A')}\n"
        f"Work Type: {_safe((permit.get('work_type') or '').replace('_', ' ').title(), 'N/A')}\n"
        f"Approval Website: {permit_url}\n"
        "\n"
        "Please review the permit, assign the permit number, and proceed it back to the receiver."
    )
    html_body = f"""
<p>A work permit has been issued and requires your <strong>HSE sign-off and permit number allocation</strong>.</p>
{_permit_info_table(permit)}
<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:13px">
  <tr style="background:#f8fbff">
    <td style="padding:10px 14px;font-weight:600;width:34%;color:#334155">Permit Issuer</td>
    <td style="padding:10px 14px;color:#0f172a">{escape(_safe(permit.get('issuer_name'), 'N/A'))}</td>
  </tr>
</table>
{_action_panel(permit, 'Open PTW For HSE Approval')}
<p style="background:#fef3c7;border-left:4px solid #f59e0b;padding:12px 16px;border-radius:6px;font-size:13px;margin-top:16px">
  <strong>Action Required - HSE:</strong> Review the permit, assign the permit number, and proceed it back to the receiver.
</p>"""
    _send(emails, subject, text, _html_wrap('HSE Sign-off Required', html_body, '#065f46'))


def notify_permit_ready_to_proceed(permit: dict, receiver_emails):
    """Step 3: Notify the receiver and maintenance engineers that work may proceed."""
    activity = _activity_details(permit)
    if isinstance(receiver_emails, str):
        emails = [receiver_emails]
    else:
        emails = list(receiver_emails or [])
    permit_url = _permit_portal_url(permit)
    subject = f"[HDEC CMMS] Permit Approved - #{_safe(permit.get('permit_number'), 'N/A')} - {_permit_subject_target(permit)}"
    text = (
        "Permit Approved\n\n"
        f"Activity: {_safe(activity.get('name'), 'N/A')}\n"
        f"Description: {_safe(activity.get('description'), 'N/A')}\n"
        f"Permit No.: {_safe(permit.get('permit_number'), 'N/A')}\n"
        f"Isolation Cert: {_safe(permit.get('isolation_cert_number'), 'N/A')}\n"
        f"Receiver: {_safe(permit.get('receiver_name'), 'N/A')}\n"
        f"Issuer: {_safe(permit.get('issuer_name'), 'N/A')}\n"
        f"HSE Officer: {_safe(permit.get('hse_name'), 'N/A')}\n"
        f"Approval Website: {permit_url}\n"
        "\n"
        "HSE has approved this permit. The receiver can now enter the permit number in CMMS and proceed the activity."
    )
    html_body = f"""
<p>HSE has approved the permit and assigned the <strong>permit number</strong>. The receiver can now proceed after confirming the same number in CMMS.</p>
{_permit_info_table(permit)}
<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:13px">
  <tr style="background:#dcfce7">
    <td style="padding:10px 14px;font-weight:700;width:34%;color:#166534">Permit Number</td>
    <td style="padding:10px 14px;font-weight:700;color:#166534">{escape(_safe(permit.get('permit_number'), 'N/A'))}</td>
  </tr>
  <tr style="background:#ffffff">
    <td style="padding:10px 14px;font-weight:600;color:#334155">Isolation Cert.</td>
    <td style="padding:10px 14px;color:#0f172a">{escape(_safe(permit.get('isolation_cert_number'), 'N/A'))}</td>
  </tr>
</table>
{_action_panel(permit, 'Open PTW To Proceed Activity')}
<p style="background:#dcfce7;border-left:4px solid #16a34a;padding:12px 16px;border-radius:6px;font-size:13px;margin-top:16px">
  <strong>Action Required:</strong> Enter the same permit number in CMMS and click proceed to unlock the activity.
</p>"""
    _send(emails, subject, text, _html_wrap('Permit Approved - Ready To Proceed', html_body, '#16a34a'))


def notify_permit_closure_requested(permit: dict, issuer_emails: list):
    """Notify the issuer that the receiver has requested permit closure."""
    activity = _activity_details(permit)
    permit_url = _permit_portal_url(permit)
    subject = f"[HDEC CMMS] Closure Required - Permit: {_permit_subject_target(permit)}"
    text = (
        "Permit Closure Requested\n\n"
        f"Activity: {_safe(activity.get('name'), 'N/A')}\n"
        f"Description: {_safe(activity.get('description'), 'N/A')}\n"
        f"Receiver: {_safe(permit.get('closure_receiver_name') or permit.get('receiver_name'), 'N/A')}\n"
        f"Issuer: {_safe(permit.get('issuer_name'), 'N/A')}\n"
        f"Permit No.: {_safe(permit.get('permit_number'), 'N/A')}\n"
        f"Approval Website: {permit_url}\n"
        "\n"
        "Please review the closure and proceed it to HSE."
    )
    html_body = f"""
<p>The receiver has requested <strong>permit closure</strong> and it now needs issuer review.</p>
{_permit_info_table(permit)}
<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:13px">
  <tr style="background:#f8fbff">
    <td style="padding:10px 14px;font-weight:600;width:34%;color:#334155">Permit Number</td>
    <td style="padding:10px 14px;color:#0f172a">{escape(_safe(permit.get('permit_number'), 'N/A'))}</td>
  </tr>
  <tr style="background:#ffffff">
    <td style="padding:10px 14px;font-weight:600;color:#334155">Closure Requested By</td>
    <td style="padding:10px 14px;color:#0f172a">{escape(_safe(permit.get('closure_receiver_name') or permit.get('receiver_name'), 'N/A'))}</td>
  </tr>
</table>
{_action_panel(permit, 'Open PTW For Closure Review')}
<p style="background:#fff7ed;border-left:4px solid #f97316;padding:12px 16px;border-radius:6px;font-size:13px;margin-top:16px">
  <strong>Action Required - Issuer:</strong> Review the closure request and send it to HSE.
</p>"""
    _send(issuer_emails, subject, text, _html_wrap('Permit Closure Requested', html_body, '#1e40af'))


def notify_permit_closure_hse_required(permit: dict, hse_officers: list):
    """Notify HSE that issuer has sent the closure for final sign-off."""
    activity = _activity_details(permit)
    permit_url = _permit_portal_url(permit)
    emails = [u.get('email', '') for u in hse_officers]
    subject = f"[HDEC CMMS] HSE Closure Required - Permit: {_permit_subject_target(permit)}"
    text = (
        "Permit Closure Awaiting HSE\n\n"
        f"Activity: {_safe(activity.get('name'), 'N/A')}\n"
        f"Description: {_safe(activity.get('description'), 'N/A')}\n"
        f"Receiver: {_safe(permit.get('receiver_name'), 'N/A')}\n"
        f"Issuer: {_safe(permit.get('closure_issuer_name') or permit.get('issuer_name'), 'N/A')}\n"
        f"Permit No.: {_safe(permit.get('permit_number'), 'N/A')}\n"
        f"Approval Website: {permit_url}\n"
        "\n"
        "Please review and close this permit."
    )
    html_body = f"""
<p>The issuer has reviewed the closure request. This permit now requires <strong>final HSE closure</strong>.</p>
{_permit_info_table(permit)}
<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:13px">
  <tr style="background:#f8fbff">
    <td style="padding:10px 14px;font-weight:600;width:34%;color:#334155">Permit Number</td>
    <td style="padding:10px 14px;color:#0f172a">{escape(_safe(permit.get('permit_number'), 'N/A'))}</td>
  </tr>
  <tr style="background:#ffffff">
    <td style="padding:10px 14px;font-weight:600;color:#334155">Closure Reviewed By</td>
    <td style="padding:10px 14px;color:#0f172a">{escape(_safe(permit.get('closure_issuer_name') or permit.get('issuer_name'), 'N/A'))}</td>
  </tr>
</table>
{_action_panel(permit, 'Open PTW For Final HSE Closure')}
<p style="background:#fef3c7;border-left:4px solid #f59e0b;padding:12px 16px;border-radius:6px;font-size:13px;margin-top:16px">
  <strong>Action Required - HSE:</strong> Complete the final closure for this permit.
</p>"""
    _send(emails, subject, text, _html_wrap('HSE Closure Required', html_body, '#065f46'))


def notify_permit_closed(permit: dict, to_emails: list):
    """Notify relevant users that the permit has been fully closed."""
    activity = _activity_details(permit)
    permit_url = _permit_portal_url(permit)
    final_pdf_line = (
        f"Final Locked PDF: {_safe(permit.get('final_pdf_url'), 'N/A')}\n"
        if permit.get('final_pdf_url') else ''
    )
    subject = f"[HDEC CMMS] Permit Closed - {_permit_subject_target(permit)}"
    text = (
        "Permit Closed\n\n"
        f"Activity: {_safe(activity.get('name'), 'N/A')}\n"
        f"Description: {_safe(activity.get('description'), 'N/A')}\n"
        f"Permit No.: {_safe(permit.get('permit_number'), 'N/A')}\n"
        f"Receiver: {_safe(permit.get('receiver_name'), 'N/A')}\n"
        f"Issuer: {_safe(permit.get('issuer_name'), 'N/A')}\n"
        f"HSE Officer: {_safe(permit.get('closure_hse_name') or permit.get('hse_name'), 'N/A')}\n"
        f"Portal: {permit_url}\n"
        f"{final_pdf_line}"
        "\n"
        "This permit is now fully closed."
    )
    html_body = f"""
<p>This <strong>Permit to Work</strong> is now fully closed.</p>
{_permit_info_table(permit)}
<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:13px">
  <tr style="background:#f8fbff">
    <td style="padding:10px 14px;font-weight:600;width:34%;color:#334155">Permit Number</td>
    <td style="padding:10px 14px;color:#0f172a">{escape(_safe(permit.get('permit_number'), 'N/A'))}</td>
  </tr>
  <tr style="background:#ffffff">
    <td style="padding:10px 14px;font-weight:600;color:#334155">Closed By HSE</td>
    <td style="padding:10px 14px;color:#0f172a">{escape(_safe(permit.get('closure_hse_name') or permit.get('hse_name'), 'N/A'))}</td>
  </tr>
</table>
{_action_panel(permit, 'Open PTW In HDEC CMMS')}
{_final_pdf_html(permit)}
<p style="background:#dcfce7;border-left:4px solid #16a34a;padding:12px 16px;border-radius:6px;font-size:13px;margin-top:16px">
  <strong>Status:</strong> Closed.
</p>"""
    _send(to_emails, subject, text, _html_wrap('Permit Closed', html_body, '#16a34a'))


def notify_activity_assigned(activity: dict, engineer_email: str, date: str = None):
    """Notify an engineer that a maintenance activity has been assigned to them."""
    if not engineer_email:
        return

    portal_url = _portal_home_url()
    subject = f"[HDEC CMMS] Maintenance Activity Assigned - {_safe(activity.get('name'), 'Activity')}"
    text = (
        "Maintenance Activity Assigned\n\n"
        f"Activity: {_safe(activity.get('name'), 'N/A')}\n"
        f"Description: {_safe(activity.get('notes'), 'N/A')}\n"
        f"Equipment: {_safe(activity.get('equipment'), 'N/A')}\n"
        f"Location: {_safe(activity.get('location'), 'N/A')}\n"
        f"Month: {_safe(activity.get('month'), 'N/A')}\n"
        + (f"Date: {date}\n" if date else "")
        + f"Portal: {portal_url}\n\n"
        + "Please log in to HDEC CMMS to view the checklist."
    )
    html_body = f"""
<p>A maintenance activity has been <strong>assigned to you</strong>.</p>
<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:13px">
  <tr style="background:#f8fbff"><td style="padding:10px 14px;font-weight:600;width:34%;color:#334155">Activity</td><td style="padding:10px 14px;color:#0f172a">{escape(_safe(activity.get('name'), 'N/A'))}</td></tr>
  <tr style="background:#ffffff"><td style="padding:10px 14px;font-weight:600;color:#334155">Description</td><td style="padding:10px 14px;color:#0f172a">{escape(_safe(activity.get('notes'), 'N/A'))}</td></tr>
  <tr style="background:#f8fbff"><td style="padding:10px 14px;font-weight:600;color:#334155">Equipment</td><td style="padding:10px 14px;color:#0f172a">{escape(_safe(activity.get('equipment'), 'N/A'))}</td></tr>
  <tr style="background:#ffffff"><td style="padding:10px 14px;font-weight:600;color:#334155">Location</td><td style="padding:10px 14px;color:#0f172a">{escape(_safe(activity.get('location'), 'N/A'))}</td></tr>
  <tr style="background:#f8fbff"><td style="padding:10px 14px;font-weight:600;color:#334155">Month</td><td style="padding:10px 14px;color:#0f172a">{escape(_safe(activity.get('month'), 'N/A'))}</td></tr>
  {f'<tr style="background:#ffffff"><td style="padding:10px 14px;font-weight:600;color:#334155">Scheduled Date</td><td style="padding:10px 14px;color:#0f172a">{escape(date)}</td></tr>' if date else ''}
</table>
<div style="margin:20px 0;padding:18px;border:1px solid #dbe4f0;border-radius:14px;background:#f8fbff">
  <div style="font-size:15px;font-weight:700;color:#0f172a;margin-bottom:6px">Open HDEC CMMS</div>
  <div style="font-size:13px;color:#475569;margin-bottom:14px">Use the website below to review the assigned activity and checklist.</div>
  <a href="{escape(portal_url)}" style="display:inline-block;background:#0f4aa3;color:#ffffff;text-decoration:none;padding:12px 18px;border-radius:10px;font-size:13px;font-weight:700">Open CMMS Portal</a>
</div>"""
    _send([engineer_email], subject, text, _html_wrap('Activity Assigned', html_body, '#1a3a6b'))


def _digest_status_label(value: str) -> str:
    status = str(value or '').strip().lower()
    mapping = {
        'completed': 'Completed',
        'done': 'Done',
        'in_progress': 'In Progress',
        'not_started': 'Not Started',
        'not_done': 'Not Done',
        'planned': 'Planned',
    }
    return mapping.get(status, status.replace('_', ' ').title() or 'Planned')


def _digest_text_section(title: str, activities: list[dict]) -> list[str]:
    lines = [title]
    if not activities:
        lines.append('No activities.')
        lines.append('')
        return lines

    for index, activity in enumerate(activities, start=1):
        lines.extend([
            f"{index}. {_safe(activity.get('name'), 'Unnamed Activity')}",
            f"   Date: {_safe(activity.get('scheduled_date'), 'N/A')}",
            f"   Status: {_digest_status_label(activity.get('status', 'planned'))}",
            f"   Description: {_safe(activity.get('notes'), 'N/A')}",
            f"   Equipment: {_safe(activity.get('equipment'), '-')}",
            f"   Location: {_safe(activity.get('location'), '-')}",
            f"   Assigned Engineer: {_safe(activity.get('assigned_engineer'), '-')}",
            f"   Assigned Technician: {_safe(activity.get('assigned_technician'), '-')}",
            "",
        ])
    return lines


def _digest_html_section(title: str, activities: list[dict], accent_color: str) -> str:
    if not activities:
        return f"""
<div style="margin:18px 0;padding:16px;border:1px solid #e2e8f0;border-radius:12px;background:#ffffff">
  <div style="font-size:15px;font-weight:700;color:{accent_color};margin-bottom:8px">{escape(title)}</div>
  <div style="font-size:13px;color:#64748b">No activities.</div>
</div>"""

    rows = []
    for activity in activities:
        rows.append(
            f"""
<tr>
  <td style="padding:9px 12px;border-top:1px solid #e5e7eb">{escape(_safe(activity.get('scheduled_date'), 'N/A'))}</td>
  <td style="padding:9px 12px;border-top:1px solid #e5e7eb">{escape(_safe(activity.get('name'), 'Unnamed Activity'))}</td>
  <td style="padding:9px 12px;border-top:1px solid #e5e7eb">{escape(_safe(activity.get('notes'), 'N/A'))}</td>
  <td style="padding:9px 12px;border-top:1px solid #e5e7eb">{escape(_safe(activity.get('equipment'), '-'))}</td>
  <td style="padding:9px 12px;border-top:1px solid #e5e7eb">{escape(_safe(activity.get('location'), '-'))}</td>
  <td style="padding:9px 12px;border-top:1px solid #e5e7eb">{escape(_safe(activity.get('assigned_engineer'), '-'))}</td>
  <td style="padding:9px 12px;border-top:1px solid #e5e7eb">{escape(_safe(activity.get('assigned_technician'), '-'))}</td>
  <td style="padding:9px 12px;border-top:1px solid #e5e7eb;font-weight:700;color:#0f172a">{escape(_digest_status_label(activity.get('status', 'planned')))}</td>
</tr>"""
        )

    return f"""
<div style="margin:18px 0;padding:16px;border:1px solid #e2e8f0;border-radius:12px;background:#ffffff">
  <div style="font-size:15px;font-weight:700;color:{accent_color};margin-bottom:10px">{escape(title)}</div>
  <table style="width:100%;border-collapse:collapse;font-size:13px">
    <thead>
      <tr style="background:#f8fbff">
        <th align="left" style="padding:9px 12px">Date</th>
        <th align="left" style="padding:9px 12px">Activity</th>
        <th align="left" style="padding:9px 12px">Description</th>
        <th align="left" style="padding:9px 12px">Equipment</th>
        <th align="left" style="padding:9px 12px">Location</th>
        <th align="left" style="padding:9px 12px">Engineer</th>
        <th align="left" style="padding:9px 12px">Technician</th>
        <th align="left" style="padding:9px 12px">Status</th>
      </tr>
    </thead>
    <tbody>{''.join(rows)}</tbody>
  </table>
</div>"""


def send_daily_activity_digest(
    digest_date: str,
    schedule_sections: dict,
    previous_24h_sections: dict,
    recipient_emails: list,
    *,
    window_end: datetime | None = None,
):
    """Send the daily maintenance schedule plus previous 24h status to maintenance engineers."""
    portal_url = _portal_home_url()
    end_dt = window_end or datetime.now()
    start_dt = end_dt - timedelta(hours=24)
    subject = f"[HDEC CMMS] Daily Maintenance Digest - {digest_date}"

    text_lines = [
        f"Daily Maintenance Digest - {digest_date}",
        "",
        f"Today's schedule date: {digest_date}",
        f"Previous 24 hours window: {start_dt.strftime('%Y-%m-%d %H:%M')} to {end_dt.strftime('%Y-%m-%d %H:%M')}",
        "",
    ]
    text_lines.extend(_digest_text_section("Today's PM Activities", schedule_sections.get('PM', [])))
    text_lines.extend(_digest_text_section("Today's CM Activities", schedule_sections.get('CM', [])))
    text_lines.extend(_digest_text_section("Previous 24 Hours - PM Activities", previous_24h_sections.get('PM', [])))
    text_lines.extend(_digest_text_section("Previous 24 Hours - CM Activities", previous_24h_sections.get('CM', [])))
    text_lines.append(f"Portal: {portal_url}")
    text = "\n".join(text_lines).rstrip()

    html_body = f"""
<p>Here is the <strong>daily maintenance digest</strong> for <strong>{escape(digest_date)}</strong>.</p>
<div style="margin:18px 0;padding:16px;border:1px solid #dbe4f0;border-radius:12px;background:#f8fbff">
  <div style="font-size:14px;font-weight:700;color:#0f172a;margin-bottom:6px">Digest Window</div>
  <div style="font-size:13px;color:#475569">Today's schedule date: <strong>{escape(digest_date)}</strong></div>
  <div style="font-size:13px;color:#475569">Previous 24 hours: <strong>{escape(start_dt.strftime('%Y-%m-%d %H:%M'))}</strong> to <strong>{escape(end_dt.strftime('%Y-%m-%d %H:%M'))}</strong></div>
</div>
{_digest_html_section("Today's PM Activities", schedule_sections.get('PM', []), '#1d4ed8')}
{_digest_html_section("Today's CM Activities", schedule_sections.get('CM', []), '#9a3412')}
{_digest_html_section("Previous 24 Hours - PM Activities", previous_24h_sections.get('PM', []), '#0f766e')}
{_digest_html_section("Previous 24 Hours - CM Activities", previous_24h_sections.get('CM', []), '#7c3aed')}
<div style="margin:20px 0;padding:18px;border:1px solid #dbe4f0;border-radius:14px;background:#f8fbff">
  <div style="font-size:15px;font-weight:700;color:#0f172a;margin-bottom:6px">Open HDEC CMMS</div>
  <div style="font-size:13px;color:#475569;margin-bottom:14px">Use the website below to review today's schedule, open activities, and permit status.</div>
  <a href="{escape(portal_url)}" style="display:inline-block;background:#0f4aa3;color:#ffffff;text-decoration:none;padding:12px 18px;border-radius:10px;font-size:13px;font-weight:700">Open CMMS Portal</a>
</div>"""

    _send(recipient_emails, subject, text, _html_wrap('Daily Maintenance Digest', html_body, '#1a3a6b'))
