import json
from datetime import datetime, timedelta
from pathlib import Path

from django.conf import settings
from django.core.management.base import BaseCommand, CommandError

from core.auth_utils import get_users_by_role
from core.cmms_utils import get_activities_for_date, get_record_for_activity_date
from core.email_utils import send_daily_activity_digest


STATE_FILE = Path(getattr(settings, 'CMMS_DATA_DIR', Path.cwd() / 'cmms_data')) / 'daily_activity_digest_state.json'


def _valid_recipients(emails: list[str]) -> list[str]:
    return list(dict.fromkeys(
        str(email).strip()
        for email in emails
        if email and '@' in str(email)
    ))


def _load_state() -> dict:
    try:
        return json.loads(STATE_FILE.read_text(encoding='utf-8'))
    except Exception:
        return {}


def _save_state(data: dict) -> None:
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps(data, indent=2), encoding='utf-8')


def _activity_type(activity: dict) -> str:
    return 'CM' if str((activity or {}).get('type', 'PM')).strip().upper() == 'CM' else 'PM'


def _activity_status(activity: dict) -> str:
    record = get_record_for_activity_date(activity.get('id', ''), activity.get('scheduled_date', ''))
    if not record:
        return 'not_started'
    if record.get('completed'):
        return 'completed'
    return 'in_progress'


def _build_schedule_sections(digest_date: str) -> dict:
    pm_items: list[dict] = []
    cm_items: list[dict] = []
    for activity in get_activities_for_date(digest_date):
        item = {
            **activity,
            'status': _activity_status(activity),
        }
        if _activity_type(activity) == 'CM':
            cm_items.append(item)
        else:
            pm_items.append(item)
    return {
        'PM': sorted(pm_items, key=lambda item: item.get('name', '')),
        'CM': sorted(cm_items, key=lambda item: item.get('name', '')),
    }


def _build_previous_24h_sections(window_end: datetime) -> dict:
    window_start = window_end - timedelta(hours=24)
    grouped = {'PM': [], 'CM': []}

    start_date = window_start.date()
    end_date = window_end.date()
    day_count = (end_date - start_date).days
    for offset in range(day_count + 1):
        current_date = (start_date + timedelta(days=offset)).isoformat()
        for activity in get_activities_for_date(current_date):
            scheduled_dt = datetime.strptime(activity.get('scheduled_date', ''), '%Y-%m-%d')
            if scheduled_dt < window_start or scheduled_dt > window_end:
                continue
            record = get_record_for_activity_date(activity.get('id', ''), activity.get('scheduled_date', ''))
            status = 'not_done'
            completed_at = ''
            started_at = ''
            if record:
                started_at = str(record.get('started_at') or '')
                completed_at = str(record.get('completed_at') or '')
                status = 'done' if record.get('completed') else 'in_progress'
            item = {
                **activity,
                'status': status,
                'started_at': started_at,
                'completed_at': completed_at,
            }
            grouped[_activity_type(activity)].append(item)

    for key in grouped:
        grouped[key] = sorted(
            grouped[key],
            key=lambda item: (item.get('scheduled_date', ''), item.get('name', '')),
        )
    return grouped


class Command(BaseCommand):
    help = "Send today's maintenance schedule and the previous 24 hours activity status to maintenance engineers by email."

    def add_arguments(self, parser):
        parser.add_argument('--date', dest='date', help='Digest date in YYYY-MM-DD format.')
        parser.add_argument('--force', action='store_true', help='Send even if this date was already sent.')
        parser.add_argument('--dry-run', action='store_true', help='Show what would be sent without sending email.')

    def handle(self, *args, **options):
        digest_date = options.get('date') or datetime.now().strftime('%Y-%m-%d')
        try:
            datetime.strptime(digest_date, '%Y-%m-%d')
        except ValueError as exc:
            raise CommandError('Date must be in YYYY-MM-DD format.') from exc

        state = _load_state()
        already_sent = state.get('last_sent_date') == digest_date
        if already_sent and not options.get('force'):
            self.stdout.write(self.style.WARNING(f'Daily activity digest already sent for {digest_date}.'))
            return

        now = datetime.now()
        schedule_sections = _build_schedule_sections(digest_date)
        previous_24h_sections = _build_previous_24h_sections(now)
        recipients = _valid_recipients([
            user.get('email', '')
            for user in get_users_by_role('maintenance_engineer')
        ])

        if not recipients:
            self.stdout.write(self.style.WARNING('No maintenance engineer email addresses are configured.'))
            return

        if options.get('dry_run'):
            self.stdout.write(f'Dry run: would send daily activity digest for {digest_date}.')
            self.stdout.write(f'Recipients: {len(recipients)}')
            self.stdout.write(
                f"Today's PM activities: {len(schedule_sections['PM'])}, "
                f"Today's CM activities: {len(schedule_sections['CM'])}"
            )
            self.stdout.write(
                f"Previous 24h PM activities: {len(previous_24h_sections['PM'])}, "
                f"Previous 24h CM activities: {len(previous_24h_sections['CM'])}"
            )
            return

        send_daily_activity_digest(
            digest_date,
            schedule_sections,
            previous_24h_sections,
            recipients,
            window_end=now,
        )
        _save_state({
            'last_sent_date': digest_date,
            'sent_at': now.isoformat(),
            'recipient_count': len(recipients),
            'activity_count': len(schedule_sections['PM']) + len(schedule_sections['CM']),
        })
        self.stdout.write(self.style.SUCCESS(
            f'Daily activity digest sent for {digest_date} to {len(recipients)} recipient(s).'
        ))
