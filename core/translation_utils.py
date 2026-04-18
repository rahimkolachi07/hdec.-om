"""
translation_utils.py - lightweight language translation helpers for chat.
"""
from __future__ import annotations

import json
import urllib.error
import urllib.parse
import urllib.request


TRANSLATION_LANGUAGES = [
    {'code': 'ar', 'label': 'Arabic', 'aliases': ['arabic']},
    {'code': 'en', 'label': 'English', 'aliases': ['english']},
    {'code': 'ur', 'label': 'Urdu', 'aliases': ['urdu']},
    {'code': 'zh-CN', 'label': 'Chinese (Simplified)', 'aliases': ['chinese', 'mandarin', 'simplified chinese', 'zh']},
    {'code': 'zh-TW', 'label': 'Chinese (Traditional)', 'aliases': ['traditional chinese']},
    {'code': 'fr', 'label': 'French', 'aliases': ['french']},
    {'code': 'es', 'label': 'Spanish', 'aliases': ['spanish']},
    {'code': 'de', 'label': 'German', 'aliases': ['german']},
    {'code': 'pt', 'label': 'Portuguese', 'aliases': ['portuguese']},
    {'code': 'tr', 'label': 'Turkish', 'aliases': ['turkish']},
    {'code': 'hi', 'label': 'Hindi', 'aliases': ['hindi']},
    {'code': 'bn', 'label': 'Bengali', 'aliases': ['bengali', 'bangla']},
    {'code': 'fa', 'label': 'Persian', 'aliases': ['persian', 'farsi']},
    {'code': 'ru', 'label': 'Russian', 'aliases': ['russian']},
    {'code': 'ja', 'label': 'Japanese', 'aliases': ['japanese']},
    {'code': 'ko', 'label': 'Korean', 'aliases': ['korean']},
    {'code': 'id', 'label': 'Indonesian', 'aliases': ['indonesian', 'bahasa']},
    {'code': 'it', 'label': 'Italian', 'aliases': ['italian']},
    {'code': 'ms', 'label': 'Malay', 'aliases': ['malay']},
    {'code': 'ta', 'label': 'Tamil', 'aliases': ['tamil']},
]


class TranslationError(Exception):
    pass


def _normalize_language(value: str) -> str:
    return ''.join(ch.lower() for ch in str(value or '').strip() if ch.isalnum())


def resolve_translation_language(value: str) -> dict | None:
    raw = str(value or '').strip()
    if not raw:
        return None
    normalized = _normalize_language(raw)
    for item in TRANSLATION_LANGUAGES:
        if raw.lower() == item['code'].lower():
            return {'code': item['code'], 'label': item['label']}
        if normalized == _normalize_language(item['label']):
            return {'code': item['code'], 'label': item['label']}
        for alias in item.get('aliases', []):
            if normalized == _normalize_language(alias):
                return {'code': item['code'], 'label': item['label']}
    return None


def extract_translation_directive(text: str) -> tuple[dict | None, str]:
    message = str(text or '').strip()
    if not message:
        return None, ''

    for prefix in ('@', '/'):
        if not message.startswith(prefix):
            continue
        token, _, remainder = message[1:].partition(' ')
        target = resolve_translation_language(token)
        if target and remainder.strip():
            return target, remainder.strip()

    return None, message


def translate_text(text: str, target_language: str) -> dict:
    target = resolve_translation_language(target_language)
    if not target:
        raise TranslationError('Unsupported translation language')

    message = str(text or '').strip()
    if not message:
        raise TranslationError('Nothing to translate')

    payload = urllib.parse.urlencode({
        'client': 'gtx',
        'sl': 'auto',
        'tl': target['code'],
        'dt': 't',
        'q': message,
    }).encode('utf-8')

    request = urllib.request.Request(
        'https://translate.googleapis.com/translate_a/single',
        data=payload,
        headers={
            'Content-Type': 'application/x-www-form-urlencoded;charset=utf-8',
            'User-Agent': 'Mozilla/5.0',
        },
    )

    try:
        with urllib.request.urlopen(request, timeout=12) as response:
            raw = response.read().decode('utf-8')
    except urllib.error.URLError as exc:
        raise TranslationError('Translation service is unavailable right now') from exc

    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise TranslationError('Translation service returned an invalid response') from exc

    translated_text = ''.join(
        part[0] for part in (parsed[0] or []) if isinstance(part, list) and part and part[0]
    ).strip()
    if not translated_text:
        raise TranslationError('Translation failed')

    source_language = ''
    if len(parsed) > 2 and isinstance(parsed[2], str):
        source_language = parsed[2]

    return {
        'text': translated_text,
        'source_language': source_language,
        'target_language': target['code'],
        'target_label': target['label'],
        'engine': 'google-translate-web',
    }
