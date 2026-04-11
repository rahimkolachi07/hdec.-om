"""
Email notification utilities for HDEC CMMS.
Sends notifications at each step of the permit workflow
and for activity assignments.
"""
import logging
from django.core.mail import send_mail
from django.conf import settings

logger = logging.getLogger(__name__)


def _send(to_emails: list, subject: str, text: str, html: str = None):
    """Send email to a list of recipients, silently ignoring failures."""
    recipients = [e for e in to_emails if e and '@' in e]
    if not recipients:
        logger.info(f"[CMMS Email] No valid recipients for: {subject}")
        return
    try:
        send_mail(
            subject=subject,
            message=text,
            from_email=getattr(settings, 'DEFAULT_FROM_EMAIL', 'HDEC CMMS <noreply@hdec.sa>'),
            recipient_list=recipients,
            html_message=html,
            fail_silently=True,
        )
        logger.info(f"[CMMS Email] Sent '{subject}' to {recipients}")
    except Exception as exc:
        logger.warning(f"[CMMS Email] Failed to send '{subject}': {exc}")


def _html_wrap(title: str, body_html: str, color: str = '#1a3a6b') -> str:
    return f"""
<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f4f6fb;font-family:Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fb;padding:30px 0">
<tr><td align="center">
<table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 16px rgba(0,0,0,.08)">
  <tr><td style="background:{color};padding:22px 30px">
    <p style="margin:0;color:#fff;font-size:11px;letter-spacing:1px;text-transform:uppercase">HDEC — 1100MW Al Henakiya Project</p>
    <h2 style="margin:6px 0 0;color:#fff;font-size:20px">{title}</h2>
  </td></tr>
  <tr><td style="padding:28px 30px;color:#2d3748;font-size:14px;line-height:1.7">
    {body_html}
    <hr style="border:none;border-top:1px solid #e2e8f0;margin:24px 0">
    <p style="font-size:12px;color:#718096">Please log in to <strong>HDEC CMMS</strong> to take action.</p>
  </td></tr>
  <tr><td style="background:#f7faff;padding:14px 30px;font-size:11px;color:#a0aec0;text-align:center">
    HDEC CMMS — Automated Notification — Do not reply to this email
  </td></tr>
</table>
</td></tr></table></body></html>"""


# ── PERMIT NOTIFICATIONS ──────────────────────────────────────────────────

def _permit_info_table(permit: dict) -> str:
    """Reusable HTML table block with permit details."""
    return f"""
<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:13px">
  <tr style="background:#f0f4ff"><td style="padding:9px 14px;font-weight:600;width:38%;color:#374151">Permit Receiver</td><td style="padding:9px 14px">{permit.get('receiver_name','—')}</td></tr>
  <tr><td style="padding:9px 14px;font-weight:600;color:#374151">Equipment</td><td style="padding:9px 14px">{permit.get('equipment','—')}</td></tr>
  <tr style="background:#f0f4ff"><td style="padding:9px 14px;font-weight:600;color:#374151">Location</td><td style="padding:9px 14px">{permit.get('location','—')}</td></tr>
  <tr><td style="padding:9px 14px;font-weight:600;color:#374151">Work Type</td><td style="padding:9px 14px">{permit.get('work_type','').replace('_',' ').title() or '—'}</td></tr>
  <tr style="background:#f0f4ff"><td style="padding:9px 14px;font-weight:600;color:#374151">Valid From</td><td style="padding:9px 14px">{permit.get('valid_from','—')}</td></tr>
  <tr><td style="padding:9px 14px;font-weight:600;color:#374151">Valid Until</td><td style="padding:9px 14px">{permit.get('valid_until','—')}</td></tr>
  <tr style="background:#f0f4ff"><td style="padding:9px 14px;font-weight:600;color:#374151">Applied On</td><td style="padding:9px 14px">{(permit.get('created_at','') or '')[:10]}</td></tr>
</table>"""


def notify_permit_created(permit: dict, operation_engineers: list, hse_officers: list = None):
    """
    Step 1: Notify operation engineers (issuers) AND HSE officers of a new permit application.
    Operation engineers must issue it; HSE get early awareness.
    """
    from datetime import datetime as _dt
    applied_date = (permit.get('created_at', '') or '')[:10]

    # ── Operation Engineers ────────────────────────────────────────────────
    op_emails = [u.get('email', '') for u in operation_engineers]
    subject_op = f"[PTW] Action Required — New Permit: {permit.get('equipment','Equipment')} | {applied_date}"
    html_op = f"""
<p style="font-size:14px">A new <strong>Permit to Work</strong> has been submitted and is awaiting your issuance.</p>
{_permit_info_table(permit)}
<p style="background:#fff7ed;border-left:4px solid #f97316;padding:12px 16px;border-radius:6px;font-size:13px">
  <strong>⚡ Action Required — Permit Issuer:</strong> Please log in to HDEC CMMS, review this permit, and issue it to proceed.
</p>"""
    _send(op_emails, subject_op,
          f"New Permit Application\nReceiver: {permit.get('receiver_name')}\nEquipment: {permit.get('equipment')}\nPlease log in to issue.",
          _html_wrap('New Permit Application — Awaiting Issuance', html_op, '#1e40af'))

    # ── HSE Officers — awareness copy ─────────────────────────────────────
    hse_list = hse_officers or []
    hse_emails = [u.get('email', '') for u in hse_list]
    if hse_emails:
        subject_hse = f"[PTW] New Permit Applied — {permit.get('equipment','Equipment')} | {applied_date}"
        html_hse = f"""
<p style="font-size:14px">A <strong>Permit to Work</strong> has been applied. This is an <em>awareness notification</em> — your sign-off will be required after the operation engineer issues the permit.</p>
{_permit_info_table(permit)}
<p style="background:#f0fdf4;border-left:4px solid #16a34a;padding:12px 16px;border-radius:6px;font-size:13px">
  ℹ️ <strong>HSE Awareness:</strong> No action required at this stage. You will receive a separate notification when HSE sign-off is needed.
</p>"""
        _send(hse_emails, subject_hse,
              f"New Permit Applied (Awareness)\nReceiver: {permit.get('receiver_name')}\nEquipment: {permit.get('equipment')}",
              _html_wrap('New Permit Applied — HSE Awareness', html_hse, '#065f46'))


def notify_permit_issued(permit: dict, hse_officers: list):
    """Step 2: Notify HSE officers permit has been issued and needs HSE sign-off."""
    emails = [u.get('email', '') for u in hse_officers]
    subject = f"[HDEC CMMS] HSE Sign-off Required — Permit: {permit.get('equipment', '')}"
    text = (
        f"Work Permit Issued — Awaiting HSE Sign-off\n\n"
        f"Receiver: {permit.get('receiver_name')}\n"
        f"Issuer: {permit.get('issuer_name')}\n"
        f"Equipment: {permit.get('equipment')}\n"
        f"Work Type: {permit.get('work_type', '').replace('_', ' ').title()}\n\n"
        f"Please log in to HDEC CMMS to assign permit/isolation numbers and sign."
    )
    html_body = f"""
<p>A Work Permit has been issued and requires your <strong>HSE sign-off and permit number allocation</strong>.</p>
<table style="width:100%;border-collapse:collapse;margin:16px 0">
  <tr style="background:#f0f4ff"><td style="padding:8px 12px;font-weight:bold;width:40%">Permit Receiver</td><td style="padding:8px 12px">{permit.get('receiver_name')}</td></tr>
  <tr><td style="padding:8px 12px;font-weight:bold">Permit Issuer</td><td style="padding:8px 12px">{permit.get('issuer_name')}</td></tr>
  <tr style="background:#f0f4ff"><td style="padding:8px 12px;font-weight:bold">Equipment</td><td style="padding:8px 12px">{permit.get('equipment')}</td></tr>
  <tr><td style="padding:8px 12px;font-weight:bold">Work Type</td><td style="padding:8px 12px">{permit.get('work_type','').replace('_',' ').title()}</td></tr>
  <tr style="background:#f0f4ff"><td style="padding:8px 12px;font-weight:bold">Valid From</td><td style="padding:8px 12px">{permit.get('valid_from')}</td></tr>
  <tr><td style="padding:8px 12px;font-weight:bold">Valid Until</td><td style="padding:8px 12px">{permit.get('valid_until')}</td></tr>
</table>
<p style="background:#fef3c7;border-left:4px solid #f59e0b;padding:10px 14px;border-radius:4px">
  🦺 <strong>Action Required:</strong> Assign permit number, isolation certificate number, and sign.
</p>"""
    _send(emails, subject, text, _html_wrap('HSE Sign-off Required', html_body, '#065f46'))


def notify_permit_approved(permit: dict, duty_engineers: list):
    """Step 3: Notify all duty maintenance engineers that permit is now active."""
    emails = [u.get('email', '') for u in duty_engineers]
    subject = (
        f"[HDEC CMMS] Permit Active — #{permit.get('permit_number')} "
        f"— {permit.get('equipment', '')}"
    )
    text = (
        f"Work Permit Now Active\n\n"
        f"Permit No.: {permit.get('permit_number')}\n"
        f"Isolation Cert: {permit.get('isolation_cert_number', 'N/A')}\n"
        f"Equipment: {permit.get('equipment')}\n"
        f"Location: {permit.get('location')}\n"
        f"Work Type: {permit.get('work_type', '').replace('_', ' ').title()}\n"
        f"Receiver: {permit.get('receiver_name')}\n"
        f"Issuer: {permit.get('issuer_name')}\n"
        f"HSE Officer: {permit.get('hse_name')}\n"
        f"Valid: {permit.get('valid_from')} to {permit.get('valid_until')}\n\n"
        f"Please ensure all safety precautions are followed."
    )
    html_body = f"""
<p>A <strong>Work Permit is now Active</strong>. All duty engineers are notified for awareness.</p>
<table style="width:100%;border-collapse:collapse;margin:16px 0">
  <tr style="background:#dcfce7"><td style="padding:8px 12px;font-weight:bold;width:40%">Permit Number</td><td style="padding:8px 12px;font-weight:bold;color:#16a34a">{permit.get('permit_number')}</td></tr>
  <tr><td style="padding:8px 12px;font-weight:bold">Isolation Cert.</td><td style="padding:8px 12px">{permit.get('isolation_cert_number','N/A')}</td></tr>
  <tr style="background:#f0f4ff"><td style="padding:8px 12px;font-weight:bold">Equipment</td><td style="padding:8px 12px">{permit.get('equipment')}</td></tr>
  <tr><td style="padding:8px 12px;font-weight:bold">Location</td><td style="padding:8px 12px">{permit.get('location')}</td></tr>
  <tr style="background:#f0f4ff"><td style="padding:8px 12px;font-weight:bold">Work Type</td><td style="padding:8px 12px">{permit.get('work_type','').replace('_',' ').title()}</td></tr>
  <tr><td style="padding:8px 12px;font-weight:bold">Receiver</td><td style="padding:8px 12px">{permit.get('receiver_name')}</td></tr>
  <tr style="background:#f0f4ff"><td style="padding:8px 12px;font-weight:bold">Issuer</td><td style="padding:8px 12px">{permit.get('issuer_name')}</td></tr>
  <tr><td style="padding:8px 12px;font-weight:bold">HSE Officer</td><td style="padding:8px 12px">{permit.get('hse_name')}</td></tr>
  <tr style="background:#f0f4ff"><td style="padding:8px 12px;font-weight:bold">Valid From</td><td style="padding:8px 12px">{permit.get('valid_from')}</td></tr>
  <tr><td style="padding:8px 12px;font-weight:bold">Valid Until</td><td style="padding:8px 12px">{permit.get('valid_until')}</td></tr>
</table>
<p style="background:#dcfce7;border-left:4px solid #16a34a;padding:10px 14px;border-radius:4px">
  ✅ Permit approved. Please ensure all isolation and safety measures are in place before work begins.
</p>"""
    _send(emails, subject, text, _html_wrap('Permit Now Active', html_body, '#16a34a'))


def notify_activity_assigned(activity: dict, engineer_email: str, date: str = None):
    """Notify an engineer that a maintenance activity has been assigned to them."""
    if not engineer_email:
        return
    subject = f"[HDEC CMMS] Maintenance Activity Assigned — {activity.get('name', '')}"
    text = (
        f"Maintenance Activity Assigned\n\n"
        f"Activity: {activity.get('name')}\n"
        f"Equipment: {activity.get('equipment')}\n"
        f"Location: {activity.get('location', '')}\n"
        f"Month: {activity.get('month')}\n"
        + (f"Date: {date}\n" if date else "")
        + "\nPlease log in to HDEC CMMS to view the checklist."
    )
    html_body = f"""
<p>A maintenance activity has been <strong>assigned to you</strong>.</p>
<table style="width:100%;border-collapse:collapse;margin:16px 0">
  <tr style="background:#f0f4ff"><td style="padding:8px 12px;font-weight:bold;width:40%">Activity</td><td style="padding:8px 12px">{activity.get('name')}</td></tr>
  <tr><td style="padding:8px 12px;font-weight:bold">Equipment</td><td style="padding:8px 12px">{activity.get('equipment')}</td></tr>
  <tr style="background:#f0f4ff"><td style="padding:8px 12px;font-weight:bold">Location</td><td style="padding:8px 12px">{activity.get('location','')}</td></tr>
  <tr><td style="padding:8px 12px;font-weight:bold">Month</td><td style="padding:8px 12px">{activity.get('month')}</td></tr>
  {f'<tr style="background:#f0f4ff"><td style="padding:8px 12px;font-weight:bold">Scheduled Date</td><td style="padding:8px 12px">{date}</td></tr>' if date else ''}
</table>"""
    _send([engineer_email], subject, text, _html_wrap('Activity Assigned', html_body, '#1a3a6b'))
