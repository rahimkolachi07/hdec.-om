import json
from datetime import datetime
from pathlib import Path

from django.conf import settings
from django.core.management.base import BaseCommand, CommandError

from core.auth_utils import get_users_by_role
from core.cmms_utils import get_activities_for_date
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


class Command(BaseCommand):
    help = "Send today's maintenance activities to all maintenance engineers by email."

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

        activities = get_activities_for_date(digest_date)
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
            self.stdout.write(f'Activities: {len(activities)}')
            return

        send_daily_activity_digest(digest_date, activities, recipients)
        _save_state({
            'last_sent_date': digest_date,
            'sent_at': datetime.now().isoformat(),
            'recipient_count': len(recipients),
            'activity_count': len(activities),
        })
        self.stdout.write(self.style.SUCCESS(
            f'Daily activity digest sent for {digest_date} to {len(recipients)} recipient(s).'
        ))
