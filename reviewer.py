"""
reviewer.py – Test Case review against SOP-033 §5.3 and TMP-10005 Rev. B.

Extracts full text from a TC .docx in document order, sends to Glean AI
for a structured compliance review, and returns a parsed result dict.
"""

import re
import json
import requests
from typing import Dict, Any, List, Optional

from docx import Document

# ---------------------------------------------------------------------------
# Embedded review checklist (SOP-033 §5.3 + TMP-10005 Rev. B)
# ---------------------------------------------------------------------------

DEFAULT_REVIEW_GUIDANCE = """\
- Flag any expected result that uses vague language such as "verify", "check", "confirm", \
"ensure", "as expected", or "screenshot taken" without defining a specific, measurable \
pass criterion.
- Flag any acceptance criterion that cannot be objectively evaluated (e.g. subjective \
wording like "appears correct" or "looks reasonable").
- If a test step evaluates a risk control measure, confirm that the expected result \
explicitly demonstrates mitigation of the associated foreseeable failure mode or hazardous \
situation — not just that the step was executed.
- Flag any test step that has a blank or placeholder expected result (e.g. "N/A", "TBD", \
or empty cells).
- Check that the Device Under Test section identifies the base part number in the format \
PN-XXX without including a version or revision number.
- Confirm that every measurement instrument listed in Section 5.1 has a unique \
calibration/traceability reference (EQ, CSV, or NPSV number).
"""

REVIEW_CHECKLIST = """\
SECTION 1 — PURPOSE & DESCRIPTION
- Describes the type/subject of testing, reason for testing, and scope.
- Requirements table present with columns: Requirement # | Description | Acceptance Criteria.
- Every listed requirement has a specific, measurable acceptance criterion (not copied from requirement text).
- Specifications table present or explicitly deleted/marked N/A.

SECTION 2 — REFERENCES
- All design input/output documents listed with document number AND revision.
- Only controlled QMS documents listed (SOPs/WIs belong in the TC SOP/WI reference, not here).

SECTION 3 — TEST INPUTS
- All test setup parameters and attributes documented.
- Each test input has objective justification that it is clinically relevant or worst-case.
- Any product modifications during testing captured in Modified Configurations table with step ref and justification.

SECTION 4 — DEVICE UNDER TEST
- Base part number identified in format PN-XXX.
- Part version/revision NOT included (belongs in the test protocol).

SECTION 5 — EQUIPMENT / TOOLS / FIXTURES / MATERIALS
- 5.1: All measurement equipment listed with unique part number/ID and traceability reference (EQ/CSV/NPSV).
- 5.2: Supporting Noah product components listed separately from DUT.
- 5.3: Non-measurement equipment/fixtures/materials listed with PLM PN or manufacturer PN, traceability, safety warnings.

SECTION 6 — TEST SETUP (required for SW test cases)
- SW build/version defined.
- OS/environment and dependencies defined.
- Setup steps table with expected results for each step.

SECTION 7 — TEST METHOD VALIDATION
- FM-023 or FM-049 assessment referenced or summarised.
- Justification provided for appropriateness of all tools/equipment.
- Section not blank.

SECTION 8 — TEST PROCEDURE
- Step-by-step procedure present.
- EVERY step has a non-blank expected result.
- Expected results are specific and measurable (not just "verify" or screenshot instructions).
- If a requirement is a risk control measure, the procedure evaluates mitigation of the foreseeable event/failure mode.

SECTION 9 — RAW DATA TABLE
- Present and formatted if test requires data collection.
- Explicitly marked N/A if workflow/pass-fail only.

SECTION 10 — DATA ANALYSIS
- Instructions present if quantitative data is collected (including ppk/distribution for variable data).
- N/A acceptable if no quantitative analysis needed.

SECTION 11 — POST-CONDITIONS / RESTORE STATE
- Expected end-state of system defined.
- Rollback/reset steps defined for any product modifications made during testing.

SECTION 14 — REVISION HISTORY
- Present and populated with at least one entry.
"""

_REVIEW_SYSTEM_PROMPT = """\
You are a senior quality engineer reviewing a software test case for compliance
with SOP-033 §5.3 (Design Verification Test Case Development) and
TMP-10005 Rev. B (Test Case Template) at a medical-device company.

Rules:
1. Be specific — cite section numbers and quote exact text when identifying issues.
2. Severity: "fail" = mandatory content missing; "warning" = present but incomplete
   or unclear; "pass" = compliant.
3. Only judge what is actually in the document — do not assume or infer missing content.
4. The document starts with a DOCUMENT SECTION MAP listing ALL headings present.
   Use this map to determine whether a section exists — do NOT mark a section as
   missing just because its body content was truncated from the extract.
5. Return ONLY a JSON object inside a ```json code block. No text outside the block.
6. NEVER include meta-commentary about the review process in findings. Findings must
   describe document quality issues only — never mention the DOCUMENT SECTION MAP,
   truncation, or the review methodology. If a section exists but its body was
   truncated, either omit a finding for that section or note only what is verifiable
   from the visible content. Do NOT write findings such as "this section appears in
   the section map" or "content was not shown in the extract".

JSON structure (strictly follow this schema):
```json
{
  "overall": "PASS" | "REVIEW REQUIRED" | "FAIL",
  "summary": "<one or two sentence overall assessment>",
  "sections": [
    {
      "id": "<e.g. '1'>",
      "name": "<section title>",
      "status": "pass" | "warning" | "fail",
      "findings": ["<specific finding 1>", "<specific finding 2>"]
    }
  ]
}
```
Only mark a section status "fail" with finding "Section not found in document." if
that section heading does NOT appear in the DOCUMENT SECTION MAP.
"""


# ---------------------------------------------------------------------------
# Text extraction — document order preserved
# ---------------------------------------------------------------------------

def extract_tc_text(docx_path: str, body_limit: Optional[int] = None) -> str:
    """
    Extract text from a TC .docx in document order (paragraphs and tables
    interleaved as they appear), preserving heading levels and table rows.

    Strategy:
    - Always prepend a SECTION MAP listing every heading found in the document
      so the AI knows which sections exist even if body content is truncated.
    - Body content is truncated to ``body_limit`` chars (caller supplies this
      so the total prompt fits within Glean's context window).
    """
    doc = Document(docx_path)
    headings: List[str] = []
    lines: List[str] = []

    def _para_text(element) -> None:
        from docx.text.paragraph import Paragraph
        p = Paragraph(element, doc)
        text = p.text.strip()
        if not text:
            return
        style = p.style.name if p.style else ''
        if any(h in style for h in ('Heading', 'Title')):
            headings.append(text)
            lines.append(f'\n### {text}')
        else:
            lines.append(text)

    def _table_text(element) -> None:
        from docx.table import Table
        tbl = Table(element, doc)
        lines.append('[TABLE]')
        for row in tbl.rows:
            seen: set = set()
            cells: List[str] = []
            for cell in row.cells:
                cid = id(cell._tc)
                if cid not in seen:
                    seen.add(cid)
                    cells.append(cell.text.strip().replace('\n', ' '))
            row_text = ' | '.join(cells)
            if row_text.strip(' |'):
                lines.append(f'  {row_text}')

    for child in doc.element.body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            _para_text(child)
        elif tag == 'tbl':
            _table_text(child)

    # Build section map — always included regardless of truncation
    section_map = ''
    if headings:
        section_map = (
            'DOCUMENT SECTION MAP (all headings found in this document):\n'
            + '\n'.join(f'  - {h}' for h in headings)
            + '\n\nFULL DOCUMENT CONTENT:\n'
        )

    body = '\n'.join(lines)
    # body_limit is passed in from review_tc so the total prompt stays within
    # Glean's context window; fall back to 12 000 if called standalone.
    actual_limit = body_limit if body_limit is not None else 12_000
    if len(body) > actual_limit:
        body = body[:actual_limit] + '\n\n[... document truncated ...]'

    return section_map + body


# ---------------------------------------------------------------------------
# Glean API call
# ---------------------------------------------------------------------------

def _call_glean(prompt: str, glean_url: str, glean_api_key: str) -> tuple:
    """
    Returns (text, error_message).  Exactly one of the two will be non-None.
    """
    endpoint = glean_url.rstrip('/') + '/rest/api/v1/chat'
    payload = {
        'messages': [{'author': 'USER', 'fragments': [{'text': prompt}]}],
        'stream': False,
    }
    headers = {
        'Authorization': f'Bearer {glean_api_key}',
        'Content-Type': 'application/json',
    }
    try:
        resp = requests.post(endpoint, json=payload, headers=headers, timeout=90)
        if not resp.ok:
            err = f'Glean returned HTTP {resp.status_code}'
            try:
                body = resp.json()
                msg = body.get('message') or body.get('error') or str(body)[:200]
                err += f': {msg}'
            except Exception:
                err += f' — {resp.text[:200]}'
            print(f'[reviewer] {err}')
            return None, err
        data = resp.json()
        for msg in reversed(data.get('messages', [])):
            if msg.get('author', '').upper() in ('GLEAN_AI', 'AI', 'ASSISTANT'):
                text = ''.join(
                    f.get('text', '') for f in msg.get('fragments', [])
                    if isinstance(f, dict)
                )
                if text.strip():
                    return text.strip(), None
        for choice in data.get('choices', []):
            content = choice.get('message', {}).get('content') or choice.get('text', '')
            if content and content.strip():
                return content.strip(), None
        return None, 'Glean returned a response but contained no text.'
    except requests.Timeout:
        err = 'Glean request timed out after 90 s — document may be too large.'
        print(f'[reviewer] {err}')
        return None, err
    except Exception as exc:
        err = f'Glean call failed: {exc}'
        print(f'[reviewer] {err}')
        return None, err


# ---------------------------------------------------------------------------
# Response parsing
# ---------------------------------------------------------------------------

def _parse_review_response(raw: str) -> Dict[str, Any]:
    """Extract JSON from Glean response."""
    m = re.search(r'```json\s*([\s\S]+?)\s*```', raw, re.IGNORECASE)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass
    m = re.search(r'\{[\s\S]+\}', raw)
    if m:
        try:
            return json.loads(m.group(0))
        except json.JSONDecodeError:
            pass
    return {
        'overall': 'REVIEW REQUIRED',
        'summary': 'AI returned an unstructured response — review manually.',
        'sections': [],
        'raw_response': raw,
    }


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def review_tc(
    docx_path: str,
    glean_config: Dict[str, str],
    extra_prompt: str = '',
) -> Dict[str, Any]:
    """
    Review a test case .docx against SOP-033 §5.3 and TMP-10005 Rev. B.

    Returns:
        {
          'overall':      'PASS' | 'REVIEW REQUIRED' | 'FAIL' | 'ERROR',
          'summary':      str,
          'sections':     [{id, name, status, findings}, ...],
          'tc_chars':     int,   # characters extracted from document
        }
    """
    extra_block = (
        f'\nAdditional reviewer guidance from user:\n{extra_prompt.strip()}\n'
        if extra_prompt and extra_prompt.strip() else ''
    )

    # Compute fixed overhead so we can budget what's left for the document.
    # Target: keep total prompt ≤ 16 000 chars (Glean's reliable working range).
    _FIXED_OVERHEAD = (
        len(_REVIEW_SYSTEM_PROMPT) + len(REVIEW_CHECKLIST) + len(extra_block)
        + 400  # formatting labels, newlines, section-map header
    )
    _MAX_TOTAL = 16_000
    body_limit = max(_MAX_TOTAL - _FIXED_OVERHEAD, 4_000)

    tc_text = extract_tc_text(docx_path, body_limit=body_limit)

    prompt = (
        f'{_REVIEW_SYSTEM_PROMPT}\n\n'
        f'REVIEW CHECKLIST (SOP-033 §5.3 + TMP-10005 Rev. B):\n{REVIEW_CHECKLIST}'
        f'{extra_block}\n'
        f'TEST CASE DOCUMENT:\n{tc_text}'
    )

    raw, glean_err = _call_glean(prompt, glean_config['url'], glean_config['api_key'])
    if not raw:
        return {
            'overall': 'ERROR',
            'summary': glean_err or 'Glean AI did not return a response. Check connection and configuration.',
            'sections': [],
            'tc_chars': len(tc_text),
        }

    result = _parse_review_response(raw)
    result['tc_chars'] = len(tc_text)
    return result
