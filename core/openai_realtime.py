"""
Helpers for OpenAI Realtime call translation sessions.

Flow:
  1. POST /v1/realtime/sessions  — create session with instructions, get ephemeral key
  2. POST /v1/realtime?model=…   — exchange SDP offer for SDP answer (Content-Type: application/sdp)
"""
from __future__ import annotations

import http.client
import json
import os
import ssl
import urllib.error
import urllib.parse
import urllib.request

from .translation_utils import resolve_translation_language


OPENAI_SESSIONS_URL = 'https://api.openai.com/v1/realtime/sessions'
OPENAI_REALTIME_URL = 'https://api.openai.com/v1/realtime'
DEFAULT_REALTIME_MODEL = 'gpt-4o-realtime-preview'
DEFAULT_REALTIME_VOICE = 'alloy'


class OpenAIRealtimeError(Exception):
    """Raised when a realtime session cannot be created."""


def _translation_prompt(target_label: str) -> str:
    return (
        'You are a low-latency live interpreter for a voice or video call. '
        f'Listen to the speaker and respond only with a faithful spoken translation in {target_label}. '
        'Do not answer questions. Do not add commentary. Do not summarize. '
        'Do not repeat the source language unless the speaker is already using the target language. '
        'Preserve tone, names, numbers, and intent as closely as possible. '
        'Keep the translation concise and natural for conversation.'
    )


def _http_post_json(url: str, body: dict, api_key: str, timeout: int = 15) -> dict:
    data = json.dumps(body).encode('utf-8')
    req = urllib.request.Request(
        url, data=data, method='POST',
        headers={
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json',
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode('utf-8'))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode('utf-8', 'replace')
        try:
            payload = json.loads(detail)
            detail = payload.get('error', {}).get('message') or detail
        except json.JSONDecodeError:
            pass
        raise OpenAIRealtimeError(detail or f'OpenAI request failed ({exc.code})') from exc
    except urllib.error.URLError as exc:
        raise OpenAIRealtimeError('OpenAI connection failed') from exc


def _http_post_sdp(url: str, sdp: str, ephemeral_key: str, timeout: int = 25) -> str:
    """POST the SDP offer using http.client for reliable body transmission."""
    parsed = urllib.parse.urlparse(url)
    sdp_bytes = sdp.encode('utf-8')
    path = parsed.path + (f'?{parsed.query}' if parsed.query else '')

    ctx = ssl.create_default_context()
    conn = http.client.HTTPSConnection(parsed.netloc, timeout=timeout, context=ctx)
    try:
        conn.request(
            'POST', path, body=sdp_bytes,
            headers={
                'Authorization': f'Bearer {ephemeral_key}',
                'Content-Type': 'application/sdp',
                'Content-Length': str(len(sdp_bytes)),
            },
        )
        resp = conn.getresponse()
        body = resp.read().decode('utf-8', 'replace')
        if resp.status != 200:
            try:
                payload = json.loads(body)
                detail = payload.get('error', {}).get('message') or body
            except json.JSONDecodeError:
                detail = body
            raise OpenAIRealtimeError(detail or f'SDP exchange failed ({resp.status})')
        return body
    except OpenAIRealtimeError:
        raise
    except Exception as exc:
        raise OpenAIRealtimeError(f'OpenAI Realtime SDP exchange failed: {exc}') from exc
    finally:
        conn.close()


def create_call_translation_session(sdp: str, target_language: str) -> tuple[str, dict]:
    api_key = os.environ.get('OPENAI_API_KEY', '').strip()
    if not api_key:
        raise OpenAIRealtimeError(
            'OPENAI_API_KEY is not configured on the server. '
            'Add it to openai_config.json.'
        )

    target = resolve_translation_language(target_language)
    if not target:
        raise OpenAIRealtimeError('Unsupported translation language')

    model = os.environ.get('OPENAI_REALTIME_MODEL', DEFAULT_REALTIME_MODEL)
    voice = os.environ.get('OPENAI_REALTIME_VOICE', DEFAULT_REALTIME_VOICE)

    # Step 1: Create a session to configure instructions and get an ephemeral key.
    session_data = _http_post_json(
        OPENAI_SESSIONS_URL,
        {
            'model': model,
            'voice': voice,
            'instructions': _translation_prompt(target['label']),
            'turn_detection': {'type': 'server_vad'},
            'input_audio_transcription': None,
            'modalities': ['audio', 'text'],
        },
        api_key,
    )

    ephemeral_key = (
        session_data.get('client_secret', {}).get('value', '').strip()
    )
    if not ephemeral_key:
        raise OpenAIRealtimeError(
            'OpenAI did not return an ephemeral key. '
            f'Response: {json.dumps(session_data)[:300]}'
        )

    # Step 2: Exchange the WebRTC SDP offer for an answer using the ephemeral key.
    sdp_url = f'{OPENAI_REALTIME_URL}?model={model}'
    answer_sdp = _http_post_sdp(sdp_url, sdp, ephemeral_key)

    return answer_sdp, target
