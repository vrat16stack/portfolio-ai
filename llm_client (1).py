"""
llm_client.py
=============
ONE gateway for every Groq call in Portfolio AI.

Fixes the problems seen in your Groq logs:
  * model shut down / renamed   -> automatic fallback through GROQ_FALLBACK_MODELS
  * 429 rate_limit_exceeded     -> paces calls by a tokens-per-minute budget and honours Retry-After
  * answers cut off at 400 tok  -> bigger token cap + reasoning_effort="low" for gpt-oss models
  * fragile "KEY: value" parse  -> JSON-schema structured output, validated before use
  * silent fake NEUTRAL verdict -> returns ok=False so callers can flag the failure

Usage:
    from llm_client import ask_json
    res = ask_json(prompt, SCHEMA, name="sentiment", validate=my_check, label="RELIANCE")
    if res.ok:   data = res.data        # dict that matches SCHEMA
    else:        print(res.error)       # every model in the chain failed

Optional lines in config.py (the YML creates them):
    GROQ_FALLBACK_MODELS = ["openai/gpt-oss-20b", "qwen/qwen3.6-27b"]
    GROQ_TPM_BUDGET      = 6500        # stay under the free-tier tokens/minute cap
"""

import json
import re
import time
from collections import deque

import config

try:
    import groq
    from groq import Groq
except ImportError:          # groq package not installed
    groq = None
    Groq = None

PRIMARY_MODEL   = getattr(config, 'GROQ_MODEL', 'openai/gpt-oss-120b')
FALLBACK_MODELS = list(getattr(config, 'GROQ_FALLBACK_MODELS', []))
TPM_BUDGET      = int(getattr(config, 'GROQ_TPM_BUDGET', 6500))
MODEL_CHAIN     = [PRIMARY_MODEL] + [m for m in FALLBACK_MODELS if m != PRIMARY_MODEL]

MIN_GAP_SECONDS   = 2.5    # keeps us under 30 requests / minute
EST_OUTPUT_TOKENS = 700    # typical size of one reply (reasoning + JSON)
MAX_429_RETRIES   = 3

STATS = {
    'calls': 0, 'ok': 0, 'failed': 0, 'retries_429': 0,
    'fallbacks': 0, 'models': {}, 'errors': [],
}

_client    = None
_windows   = {}    # model -> deque[(timestamp, tokens)] for the last 60 s
_last_call = {}    # model -> timestamp of last request
_mode      = {}    # model -> 'schema' | 'json_object' | 'plain'  (what the model accepts)


class LLMResult:
    def __init__(self, ok, data=None, model='', error='', raw=''):
        self.ok    = ok
        self.data  = data
        self.model = model
        self.error = error
        self.raw   = raw


def get_stats():
    return dict(STATS, models=dict(STATS['models']), errors=list(STATS['errors']))


def reset_stats():
    STATS.update({'calls': 0, 'ok': 0, 'failed': 0, 'retries_429': 0,
                  'fallbacks': 0, 'models': {}, 'errors': []})


def _get_client():
    global _client
    if Groq is None:
        raise RuntimeError("groq package is not installed")
    if _client is None:
        _client = Groq(api_key=config.GROQ_API_KEY, max_retries=0, timeout=60)
    return _client


# ── helpers ───────────────────────────────────────────────────────────────────

def _schema_hint(schema):
    parts = []
    for key, spec in schema.get('properties', {}).items():
        if 'enum' in spec:
            parts.append(f'"{key}": one of {spec["enum"]}')
        else:
            parts.append(f'"{key}": {spec.get("type", "string")}')
    return ("\n\nReturn ONLY one JSON object (no markdown, no extra text) "
            "with exactly these keys:\n" + "\n".join(parts))


def _parse_json(text):
    if not text:
        return None
    t = re.sub(r'^```(?:json)?\s*|\s*```$', '', text.strip())
    try:
        return json.loads(t)
    except Exception:
        pass
    m = re.search(r'\{.*\}', t, re.S)
    if m:
        try:
            return json.loads(m.group(0))
        except Exception:
            return None
    return None


def _classify(e):
    text   = str(e).lower()
    status = getattr(e, 'status_code', None)
    if status == 429 or (groq is not None and isinstance(e, groq.RateLimitError)):
        if any(k in text for k in ('per day', 'tokens per day', 'requests per day')):
            return 'daily_limit'
        return 'rate_limit'
    if status in (400, 404):
        if 'model' in text and any(k in text for k in (
                'decommission', 'model_not_found', 'does not exist',
                'no longer supported', 'deprecat', 'not found')):
            return 'model_gone'
        if any(k in text for k in ('response_format', 'json_schema', 'json mode',
                                   'structured output', 'schema')):
            return 'schema_unsupported'
    if status in (401, 403):
        return 'auth'
    if status is not None and status >= 500:
        return 'transient'
    if groq is not None and isinstance(e, (groq.APIConnectionError, groq.APITimeoutError)):
        return 'transient'
    return 'other'


def _retry_after(e, default=15.0):
    try:
        header = e.response.headers.get('retry-after')
        if header:
            return float(header)
    except Exception:
        pass
    m = re.search(r'try again in\s*(?:(\d+)m)?\s*([\d.]+)s', str(e))
    if m:
        return int(m.group(1) or 0) * 60 + float(m.group(2))
    return default


def _note_usage(model, tokens):
    _windows.setdefault(model, deque()).append((time.time(), int(tokens)))


def _throttle(model, est_tokens):
    """Wait until sending est_tokens keeps this model under the per-minute token budget."""
    est_tokens = min(est_tokens, TPM_BUDGET)
    window = _windows.setdefault(model, deque())
    last   = _last_call.get(model)
    if last is not None:
        gap = MIN_GAP_SECONDS - (time.time() - last)
        if gap > 0:
            time.sleep(gap)
    while True:
        now = time.time()
        while window and now - window[0][0] > 60:
            window.popleft()
        used = sum(t for _, t in window)
        if used + est_tokens <= TPM_BUDGET or not window:
            return
        time.sleep(max(60 - (now - window[0][0]) + 0.5, 0.5))


def _build_kwargs(model, prompt, schema, name, mode, max_tokens, temperature):
    content = prompt if mode == 'schema' else prompt + _schema_hint(schema)
    kwargs = {
        'model':       model,
        'messages':    [{'role': 'user', 'content': content}],
        'max_tokens':  max_tokens,
        'temperature': temperature,
    }
    if mode == 'schema':
        kwargs['response_format'] = {
            'type': 'json_schema',
            'json_schema': {'name': name, 'strict': True, 'schema': schema},
        }
    elif mode == 'json_object':
        kwargs['response_format'] = {'type': 'json_object'}
    if model.startswith('openai/gpt-oss'):
        kwargs['extra_body'] = {'reasoning_effort': 'low'}   # less hidden thinking, no truncation
    return kwargs


# ── one model, with retries ───────────────────────────────────────────────────

def _try_model(model, prompt, schema, name, validate, max_tokens, temperature, errors):
    mode      = _mode.get(model, 'schema')
    cap       = max_tokens
    rl_tries  = 0
    net_tries = 0
    grew      = False
    reparsed  = False

    while True:
        kwargs = _build_kwargs(model, prompt, schema, name, mode, cap, temperature)
        est    = len(kwargs['messages'][0]['content']) // 3 + 60 + EST_OUTPUT_TOKENS
        _throttle(model, est)

        try:
            resp = _get_client().chat.completions.create(**kwargs)
        except Exception as e:
            kind = _classify(e)
            if kind == 'rate_limit' and rl_tries < MAX_429_RETRIES:
                rl_tries += 1
                STATS['retries_429'] += 1
                _note_usage(model, min(est, TPM_BUDGET))
                time.sleep(min(_retry_after(e) + 1.0, 60.0))
                continue
            if kind == 'transient' and net_tries < 2:
                net_tries += 1
                time.sleep(4 * net_tries)
                continue
            if kind == 'schema_unsupported' and mode != 'plain':
                mode = 'json_object' if mode == 'schema' else 'plain'
                _mode[model] = mode
                continue
            errors.append(f'{model}: {kind} - {str(e)[:140]}')
            return None

        _last_call[model] = time.time()
        choice  = resp.choices[0]
        content = (choice.message.content or '').strip()
        usage   = getattr(resp, 'usage', None)
        _note_usage(model, getattr(usage, 'total_tokens', None) or est)
        finish  = getattr(choice, 'finish_reason', '')

        data = _parse_json(content)
        if data is None:
            if finish == 'length' and not grew:       # cut off -> allow a longer answer
                grew = True
                cap  = min(cap * 2, 3000)
                continue
            if not reparsed:
                reparsed = True
                continue
            errors.append(f'{model}: empty or unparseable reply (finish={finish or "?"})')
            return None

        if validate is not None:
            try:
                valid = bool(validate(data))
            except Exception:
                valid = False
            if not valid:
                errors.append(f'{model}: reply failed validation')
                return None

        return data, content


# ── public entry point ────────────────────────────────────────────────────────

def ask_json(prompt, schema, name='response', validate=None,
             max_tokens=1200, temperature=0.2, label=''):
    """Ask the model chain for ONE JSON object. Never raises; check .ok."""
    STATS['calls'] += 1
    errors = []
    for idx, model in enumerate(MODEL_CHAIN):
        out = _try_model(model, prompt, schema, name, validate,
                         max_tokens, temperature, errors)
        if out is not None:
            data, raw = out
            STATS['ok'] += 1
            STATS['models'][model] = STATS['models'].get(model, 0) + 1
            if idx > 0:
                STATS['fallbacks'] += 1
                print(f"[llm] {label}: primary model failed, answered by fallback {model}")
            return LLMResult(True, data, model, '', raw)

    STATS['failed'] += 1
    message = ' | '.join(errors[-3:]) or 'all models failed'
    STATS['errors'].append(f'{label}: {message}'[:300])
    print(f"[llm] {label}: ALL models failed -> {message}")
    return LLMResult(False, None, '', message, '')
