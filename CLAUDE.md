# Test Case Toolbox — Claude Code Context

## What this app does
Flask web app (single-user, local) that parses `.docx` software test-case
documents, maps each Requirement / Specification to its matched test procedure
steps, and generates compressed statement-style Acceptance Criteria using the
**Glean Chat API** (with a rule-based fallback when AI is not configured).

## How to run / preview

### One-command dev start (Claude Code preview)
```
npm run dev
```
This single command starts **both** servers automatically:
- Flask backend on **port 5000** (via the `vite-plugin-flask` in `vite.config.ts`)
- Vite dev server on **port 5173** (proxies all requests to Flask)

**Preview URL: http://localhost:5173**

### Prerequisites (one-time setup)
```
deploy.bat   # creates venv/ and installs Python deps
npm install  # installs Vite + TypeScript
```

### Other commands
```
stop.bat     # kill Flask on port 5000
npm run build   # compile frontend assets → static/dist/
```

## File map
| File | Purpose |
|---|---|
| `app.py` | Flask app — all routes |
| `parser.py` | `.docx` parser — extracts Requirements, Specifications, Procedure Steps |
| `summarizer.py` | Glean AI client + rule-based fallback AC generator |
| `reviewer.py` | TC compliance review against SOP-033 §5.3 + TMP-10005 Rev. B |
| `exporter.py` | Excel (openpyxl) and Word (python-docx) export |
| `templates/index.html` | Entire front-end — three tabs: AC Builder, TC Review, Setup |
| `requirements.txt` | `flask`, `python-docx`, `openpyxl`, `requests` |
| `deploy.bat` | Auto-discovers Python ≥ 3.10, creates `venv/`, installs deps |
| `launch.bat` | Starts `venv/Scripts/python.exe app.py`, waits for port 5000, opens browser |
| `glean_config.json` | **Gitignored** — persisted Glean URL + API key |
| `setup_config.json` | **Gitignored** — persisted AC prompt + review extra guidance |
| `references/` | SOP-033 Rev. I and TMP-10005 Rev. B source docs (and any user-added refs) |
| `vite.config.ts` | Vite build config + `vite-plugin-flask` (auto-starts Flask on `npm run dev`) |
| `package.json` | Node dev deps: `vite`, `typescript` |
| `frontend/index.html` | Vite entry point (dev only — Flask/Jinja2 renders the real template) |
| `frontend/main.ts` | JS bundle entry — grows as inline JS is extracted from `templates/index.html` |
| `static/dist/` | Vite build output — referenced by Flask template in production |

## Architecture

### Parser (`parser.py`)
- Identifies tables by **column headers**, not section numbers (robust to document reformatting)
- `_is_requirement_table()` — looks for "requirement #" in headers
- `_is_specification_table()` — looks for "specification #" in headers
- `_is_procedure_table()` — requires "procedure step description" (strict, avoids grabbing the Pre-Test Setup table)
- `numbers_match(a, b)` — suffix comparison so "UIS-112369" matches "TBR-UIS-112369"
- `REQ_NUMBER_RE` — matches `(?:TBR-)?(?:SRS|UIS|SDS)-\d{4,6}`
- Returns `ParseResult(requirements, specifications, procedure_steps, errors, warnings)`

### Summarizer (`summarizer.py`)
- **Primary path**: `POST {glean_url}/rest/api/v1/chat` — Bearer auth, `messages[{author:"USER", fragments:[{text:prompt}]}]`, `stream:false`
- **Response parsing**: looks for `author == GLEAN_AI / AI / ASSISTANT` in `messages[]`, falls back to OpenAI `choices[]` format
- **Prompt design**: requirement text is labelled "Context — DO NOT copy"; test steps are the "Primary source". AI is instructed to produce ONE sentence derived from test evidence, not the requirement wording.
- `custom_prompt` parameter threads from `/regenerate-single` → `generate_acceptance_criteria_for` → `_glean_synthesise` → `_build_user_message`
- `_clean_ai_text()` strips `---` separators and "Would you like…" follow-ups
- **Fallback**: `_rule_based_ac()` — cleans expected results, deduplicates, joins as prose paragraph
- `glean_config.json` is read on startup and written on every `POST /configure` call

### Key Flask routes (`app.py`)
| Route | Method | Purpose |
|---|---|---|
| `/upload` | POST | Parse doc, run `generate_all()`, cache results in `_state` |
| `/regenerate` | POST | Re-run AI on all cached items |
| `/regenerate-single` | POST | Re-run AI on one item `{number, custom_prompt}` |
| `/save-edit` | POST | Persist manually-edited AC text `{number, text}` to `_state` |
| `/configure` | POST | Save `{url, api_key}` to `_state` + `glean_config.json` |
| `/test-connection` | POST | Sends test call to Glean, returns `{ok, message}` |
| `/ai-status` | GET | Returns `{ai_active, glean_url, api_key, has_key, message}` |
| `/export/excel` | GET | Download `.xlsx` from `_state['results']` |
| `/export/docx` | GET | Download `.docx` from `_state['results']` |
| `/review-tc` | POST | Upload TC .docx → run `reviewer.review_tc()` → return structured result |
| `/api/setup-config` | GET | Return `{ac_system_prompt, review_extra_guidance, default_ac_prompt}` |
| `/api/save-setup` | POST | Persist `{ac_system_prompt, review_extra_guidance}` to `setup_config.json` |
| `/api/references` | GET | List files in `references/` folder |
| `/api/upload-reference` | POST | Upload a new doc to `references/` |

### Front-end (`templates/index.html`)
- **Single HTML file** — inline CSS + vanilla JS, no build step, no frameworks
- `_allResults` — JS mirror of the server's `_state['results']`; updated on upload, regenerate, and per-card generate
- **AC textarea**: starts `readonly`; **Edit** button → editable + blue border; **Save** button → `POST /save-edit` → back to readonly
- **Per-card Generate with AI**: amber button → inline prompt panel → `POST /regenerate-single` → updates textarea + source badge
- **Export**: `triggerExport(format)` uses `fetch()` + `URL.createObjectURL()` for file download (not `<a href>` navigation)
- **Glean config modal**: pre-populated from `/ai-status` on open; Test Connection button; saves to `/configure`

## State management
`_state` in `app.py` is an in-memory dict (single-user local app — no sessions/DB).
Results are lost on server restart; users must re-upload. Config persists via `glean_config.json`.

## Known constraints / design decisions
- `parser.py` imports shadow the stdlib `parser` module — do not rename to conflict
- The app uses `from parser import ...` — keep `parser.py` in the same directory as `app.py`
- `venv/` and `glean_config.json` are gitignored; both are created locally at runtime
- Export reads from `_state['results']` (server-side); manual edits via `/save-edit` keep it in sync
- No authentication — designed for localhost use only
