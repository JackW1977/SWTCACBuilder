"""
summarizer.py – Acceptance Criteria generator.

Primary path  : Glean Chat API  –  POST {base_url}/rest/api/v1/chat
                Configured at runtime via /configure; no SDK required.

Fallback path : Rule-based extraction used when Glean config is absent or
                if the API call fails.

The AI prompt explicitly instructs the model to synthesise—not copy—the
expected results, and to write a single cohesive paragraph in "shall" language.
"""

import re
import textwrap
import requests
from typing import List, Dict, Any, Optional

from parser import RequirementRow, ProcedureStep, numbers_match


# ---------------------------------------------------------------------------
# Prompt templates
# ---------------------------------------------------------------------------

_SYSTEM_INSTRUCTIONS = textwrap.dedent("""\
    You are a senior technical writer for medical-device software verification.
    Write ONE acceptance-criteria (AC) statement for the given requirement.

    CRITICAL rules — follow in order of priority:
    1. Derive the AC from the TEST EVIDENCE (steps + expected results), NOT from
       the requirement text. The requirement tells you WHAT is being tested;
       the test steps tell you HOW it is verified — write the AC from the HOW.
    2. Do NOT copy, paraphrase, or mirror the requirement wording. The AC must
       read as independent verification language, not a restatement.
    3. Use formal "shall" language and be maximally concise: ONE sentence.
       Two sentences only if critical information genuinely cannot be compressed.
    4. Combine all test conditions into one tight assertion using connectors
       ("and", "while", "upon", "when", "within"). No lists, no bullets,
       no line breaks between items.
    5. Strip ALL procedural noise: "take a screenshot", "verify that",
       "attach to results", "compare with reference", "mark as PASS/FAIL",
       "user clicks", "navigate to", etc.
    6. Preserve exact technical identifiers and capitalisation verbatim
       (e.g. Treatment_Delivery_Software, Robot_Arm,
       Metal-Suppressed_3D_Image_Volume).
    7. Output ONLY the AC text — no ID prefix, no step numbers, no commentary.
""")


def _build_user_message(req_number: str, req_description: str,
                        steps: List[ProcedureStep],
                        custom_prompt: str = '') -> str:
    steps_block = ""
    for step in steps:
        clean_er = re.sub(
            r'[Tt]ake a screenshot[^.\n]*[.\n]?|'
            r'[Aa]ttach[^.\n]*(results|evidence)[^.\n]*[.\n]?|'
            r'[Cc]ompare with the reference[^.\n]*[.\n]?|'
            r'[Mm]ark this step as (PASS|FAIL)[^.\n]*[.\n]?',
            '', step.expected_result,
        ).strip()
        steps_block += (
            f"\nStep {step.step_id}:\n"
            f"  Description    : {step.description[:400]}\n"
            f"  Expected Result: {clean_er[:400]}\n"
        )

    custom_block = (
        f"\nAdditional instruction from user: {custom_prompt.strip()}\n"
        if custom_prompt and custom_prompt.strip() else ""
    )

    return textwrap.dedent(f"""\
        {_SYSTEM_INSTRUCTIONS}

        Context (DO NOT copy this wording into the AC):
        Requirement ID  : {req_number}
        Requirement text: {req_description}

        Primary source — derive the AC from this test evidence:
        {steps_block}{custom_block}
        Write the acceptance criteria. Base it on the test evidence above,
        not on the requirement text.
    """)


# ---------------------------------------------------------------------------
# Glean HTTP client
# ---------------------------------------------------------------------------

def _glean_synthesise(
    req_number: str,
    req_description: str,
    steps: List[ProcedureStep],
    glean_url: str,
    glean_api_key: str,
    custom_prompt: str = '',
) -> Optional[str]:
    """
    Call the Glean Chat API and return the generated AC text.
    Returns None on any failure so the caller falls back to rule-based.

    Endpoint  : POST {glean_url}/rest/api/v1/chat
    Auth      : Authorization: Bearer {glean_api_key}
    """
    endpoint = glean_url.rstrip('/') + '/rest/api/v1/chat'
    prompt   = _build_user_message(req_number, req_description, steps, custom_prompt)

    payload = {
        "messages": [
            {
                "author":    "USER",
                "fragments": [{"text": prompt}],
            }
        ],
        "stream": False,
    }

    headers = {
        "Authorization": f"Bearer {glean_api_key}",
        "Content-Type":  "application/json",
    }

    try:
        resp = requests.post(endpoint, json=payload, headers=headers, timeout=45)
        resp.raise_for_status()
        data = resp.json()

        # --- Parse response: Glean-native format ---
        for msg in reversed(data.get("messages", [])):
            if msg.get("author", "").upper() in ("GLEAN_AI", "AI", "ASSISTANT"):
                text = "".join(
                    f.get("text", "") for f in msg.get("fragments", [])
                    if isinstance(f, dict)
                )
                if text.strip():
                    return _clean_ai_text(text)

        # --- Fallback: OpenAI-compatible format ---
        for choice in data.get("choices", []):
            content = (
                choice.get("message", {}).get("content")
                or choice.get("text", "")
            )
            if content and content.strip():
                return _clean_ai_text(content)

        # --- Last resort: first non-empty string value in response ---
        raw = str(data)
        if len(raw) > 20:
            print(f"[summarizer] Unexpected Glean response shape for {req_number}: {raw[:200]}")

        return None

    except requests.exceptions.HTTPError as exc:
        print(f"[summarizer] Glean HTTP error for {req_number}: {exc} — {exc.response.text[:200]}")
        return None
    except Exception as exc:
        print(f"[summarizer] Glean call failed for {req_number}: {exc}")
        return None


_AI_NOISE_RE = re.compile(
    r'\n\s*---.*$|'                                   # --- separator and everything after
    r'\s*Would you like\b[^.?!]*[.?!]?|'             # "Would you like..." follow-ups
    r'\s*Do you (?:need|want|require)\b[^.?!]*[.?!]?|'
    r'\s*(?:Is there|Are there) anything\b[^.?!]*[.?!]?|'
    r'\s*(?:Let me know|Feel free)\b[^.?!]*[.?!]?|'
    r'\s*(?:Note|Please note)[:\s][^.]*\.',           # "Note: ..." trailing notes
    re.IGNORECASE | re.DOTALL,
)


def _clean_ai_text(text: str) -> str:
    """Strip follow-up recommendations and conversational artefacts from AI output."""
    text = _AI_NOISE_RE.sub('', text)
    # Collapse multiple blank lines left by stripping
    text = re.sub(r'\n{2,}', '\n', text)
    return text.strip()


def test_glean_connection(glean_url: str, glean_api_key: str) -> Dict[str, Any]:
    """
    Send a minimal test message to Glean and return {ok, message}.
    Used by the /test-connection endpoint.
    """
    result = _glean_synthesise(
        req_number="TEST",
        req_description="The software shall respond to a test query.",
        steps=[],
        glean_url=glean_url,
        glean_api_key=glean_api_key,
    )
    if result is not None:
        return {"ok": True,  "message": "Connection successful."}
    return {"ok": False, "message": "No valid response from Glean. Check URL and API key."}


# ---------------------------------------------------------------------------
# Rule-based fallback
# ---------------------------------------------------------------------------

_STRIP_RE = re.compile(
    r'[Tt]ake a screenshot[^.]*\.|'
    r'[Aa]ttach (it|the screenshot|a screenshot|the (image|evidence))[^.]*\.|'
    r'[Aa]ttach[^.]*results column[^.]*\.|'
    r'[Cc]ompare with the reference[^.]*\.|'
    r'[Cc]ompare to the reference[^.]*\.|'
    r'[Ss]creenshot[s]? (is |are )?attached[^.]*\.|'
    r'[Mm]ark this step as (PASS|FAIL)[^.]*\.|'
    r'[Ss]ee reference image[s]?[^.]*\.|'
    r'[Rr]eference image[s]? below[^.]*\.|'
    r'as an objective evidence[^.]*\.|'
    r'in the results column[^.]*\.|'
    r'[Mm]entioned actions are performed[^.]*\.'
)

_PROCEDURAL_LINE_RE = re.compile(
    r'^(User (is|can|is able to|shall|has|drives|waits|proceeds)|'
    r'Procedure [A-C] (is|has|was)|'
    r'System is (shutdown|initialized|turned|powered)|'
    r'Mentioned actions are performed|'
    r'[A-Z][a-z]+ (is|are) (attached|exported|completed|performed|shutdown'
    r'|powered|extracted|transferred|signed|launched|recorded|started|captured|ready)|'
    r'N/A$)',
    re.IGNORECASE,
)


def _rule_based_ac(steps: List[ProcedureStep]) -> str:
    parts: List[str] = []
    for step in steps:
        for raw in step.expected_result.splitlines():
            line = _STRIP_RE.sub('', raw.strip()).strip()
            line = re.sub(r'^\s*[Vv]erify\s+(?:that\s+)?', '', line).strip()
            line = re.sub(r'\s*[Tt]ake a screenshot.*$', '', line).rstrip('. ')
            if not line or len(line) < 10:
                continue
            if _PROCEDURAL_LINE_RE.match(line):
                continue
            parts.append(line)
        for note_line in step.notes.splitlines():
            note_line = note_line.strip()
            if 'shall' in note_line.lower() and len(note_line) > 20:
                parts.append(note_line)

    seen: set = set()
    unique: List[str] = []
    for p in parts:
        key = re.sub(r'\s+', ' ', p.lower())
        if key not in seen:
            seen.add(key)
            unique.append(p)

    if not unique:
        return ""

    normalised = []
    for t in unique:
        t = t[0].upper() + t[1:] if t else t
        if t and t[-1] not in '.!?':
            t += '.'
        normalised.append(t)

    return ' '.join(normalised)


# ---------------------------------------------------------------------------
# Core per-item generation
# ---------------------------------------------------------------------------

def generate_acceptance_criteria_for(
    item: RequirementRow,
    all_steps: List[ProcedureStep],
    glean_config: Optional[Dict] = None,
    custom_prompt: str = '',
) -> Dict[str, Any]:
    matching: List[ProcedureStep] = [
        s for s in all_steps
        if any(numbers_match(item.number, ref) for ref in s.referenced_numbers)
    ]

    base = {
        'number':       item.number,
        'description':  item.description,
        'existing_ac':  item.existing_ac,
        'matching_steps': [
            {
                'step_id':        s.step_id,
                'description':    s.description[:350],
                'expected_result': s.expected_result[:450],
            }
            for s in matching
        ],
        'step_count':   len(matching),
        'unmatched':    len(matching) == 0,
        'ai_generated': False,
    }

    if not matching:
        base['generated_ac'] = ''
        return base

    ac_text: Optional[str] = None

    # Try Glean AI if config is provided
    if glean_config and glean_config.get('url') and glean_config.get('api_key'):
        ac_text = _glean_synthesise(
            item.number,
            item.description,
            matching,
            glean_config['url'],
            glean_config['api_key'],
            custom_prompt=custom_prompt,
        )
        if ac_text:
            base['ai_generated'] = True

    if not ac_text:
        ac_text = _rule_based_ac(matching)

    base['generated_ac'] = ac_text or ''
    return base


# ---------------------------------------------------------------------------
# Public entry points
# ---------------------------------------------------------------------------

def generate_all(
    requirements: List[RequirementRow],
    specifications: List[RequirementRow],
    procedure_steps: List[ProcedureStep],
    use_ai: bool = True,
    glean_config: Optional[Dict] = None,
) -> Dict[str, Any]:
    effective_config = glean_config if use_ai else None
    results: Dict[str, Any] = {}

    for item in requirements:
        entry = generate_acceptance_criteria_for(item, procedure_steps, effective_config)
        entry['item_type'] = 'Requirement'
        results[item.number] = entry

    for item in specifications:
        entry = generate_acceptance_criteria_for(item, procedure_steps, effective_config)
        entry['item_type'] = 'Specification'
        results[item.number] = entry

    return results


def ai_available(glean_config: Optional[Dict] = None) -> bool:
    """Return True if a valid Glean config (url + api_key) is stored."""
    if not glean_config:
        return False
    return bool(glean_config.get('url') and glean_config.get('api_key'))
