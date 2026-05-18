"""
app.py – Flask web application for the SW Test Case Acceptance Criteria Builder.

Routes
------
GET  /                        – main UI
POST /upload                  – parse .docx, return JSON results
POST /regenerate              – re-run AC generation on cached data (use_ai=True)
POST /regenerate-single       – re-run AC for one item with optional custom prompt
POST /save-edit               – persist manually-edited AC to session state
POST /configure               – save Glean service URL + API key
POST /test-connection         – test Glean connectivity
GET  /ai-status               – report whether Glean AI is configured
GET  /export/excel            – download .xlsx
GET  /export/docx             – download .docx
GET  /health                  – liveness check
POST /review-tc               – review a TC .docx against SOP-033 §5.3 + TMP-10005
GET  /api/setup-config        – return setup config + default AC prompt
POST /api/save-setup          – persist AC prompt + review guidance
GET  /api/references          – list files in references/ folder
POST /api/upload-reference    – upload a new file to references/ folder
"""

import io
import os
import json
import traceback
import tempfile

from flask import Flask, render_template, request, jsonify, send_file

from parser import parse_document
from summarizer import (generate_all, generate_acceptance_criteria_for,
                        ai_available, test_glean_connection,
                        set_ac_system_prompt, DEFAULT_AC_SYSTEM_PROMPT)
from exporter import export_to_excel, export_to_docx
import reviewer

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.secret_key = os.urandom(32)

# ---------------------------------------------------------------------------
# Persisted Glean config
# ---------------------------------------------------------------------------
_CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'glean_config.json')
_SETUP_CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'setup_config.json')
_REFERENCES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'references')


def _load_glean_config() -> dict:
    """Load saved Glean config from disk, fall back to env vars."""
    try:
        with open(_CONFIG_FILE, 'r') as f:
            cfg = json.load(f)
            if cfg.get('url') and cfg.get('api_key'):
                return cfg
    except Exception:
        pass
    return {
        'url':     os.environ.get('GLEAN_URL', ''),
        'api_key': os.environ.get('GLEAN_API_KEY', ''),
    }


def _save_glean_config(cfg: dict) -> None:
    try:
        with open(_CONFIG_FILE, 'w') as f:
            json.dump(cfg, f)
    except Exception as exc:
        print(f'[config] Could not persist Glean config: {exc}')


def _load_setup_config() -> dict:
    """Load setup config (AC prompt + review guidance) from disk."""
    try:
        with open(_SETUP_CONFIG_FILE, 'r') as f:
            return json.load(f)
    except Exception:
        pass
    return {'ac_system_prompt': '', 'review_extra_guidance': ''}


def _save_setup_config(cfg: dict) -> None:
    try:
        with open(_SETUP_CONFIG_FILE, 'w') as f:
            json.dump(cfg, f, indent=2)
    except Exception as exc:
        print(f'[config] Could not persist setup config: {exc}')


# ---------------------------------------------------------------------------
# In-memory session state
# ---------------------------------------------------------------------------
def _init_state() -> dict:
    glean_cfg = _load_glean_config()
    setup_cfg = _load_setup_config()
    # Apply saved AC prompt if one exists
    if setup_cfg.get('ac_system_prompt'):
        set_ac_system_prompt(setup_cfg['ac_system_prompt'])
    return {
        'results':          {},
        'filename':         '',
        'stats':            {},
        'warnings':         [],
        '_requirements':    [],
        '_specifications':  [],
        '_procedure_steps': [],
        'glean_config':     glean_cfg,
        'setup_config':     setup_cfg,
    }

_state: dict = _init_state()


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/health')
def health():
    return jsonify({'status': 'ok'})


@app.route('/ai-status')
def ai_status():
    active = ai_available(_state['glean_config'])
    cfg    = _state['glean_config']
    return jsonify({
        'ai_active':  active,
        'glean_url':  cfg.get('url', ''),
        'api_key':    cfg.get('api_key', ''),
        'has_key':    bool(cfg.get('api_key')),
        'message': 'Glean AI synthesis enabled.' if active
                   else 'Glean not configured — using rule-based fallback.',
    })


@app.route('/configure', methods=['POST'])
def configure():
    """Save Glean service URL and API key."""
    data    = request.get_json(silent=True) or {}
    url     = (data.get('url')     or '').strip().rstrip('/')
    api_key = (data.get('api_key') or '').strip()

    if not url:
        return jsonify({'success': False, 'error': 'Glean service URL is required.'}), 400
    if not api_key:
        return jsonify({'success': False, 'error': 'API key is required.'}), 400
    if not url.startswith('http'):
        return jsonify({'success': False, 'error': 'URL must start with http:// or https://'}), 400

    cfg = {'url': url, 'api_key': api_key}
    _state['glean_config'] = cfg
    _save_glean_config(cfg)
    return jsonify({'success': True, 'message': 'Glean configuration saved.'})


@app.route('/save-edit', methods=['POST'])
def save_edit():
    """Persist a manually-edited AC text back to session state (for export)."""
    data = request.get_json(silent=True) or {}
    number = (data.get('number') or '').strip()
    text   = data.get('text', '')
    if not number or number not in _state['results']:
        return jsonify({'ok': False}), 400
    _state['results'][number]['generated_ac'] = text
    return jsonify({'ok': True})


@app.route('/test-connection', methods=['POST'])
def test_connection():
    """Test connectivity to the configured Glean endpoint."""
    cfg = _state['glean_config']
    if not ai_available(cfg):
        return jsonify({'ok': False, 'message': 'No Glean config saved yet.'}), 400
    result = test_glean_connection(cfg['url'], cfg['api_key'])
    return jsonify(result), 200 if result['ok'] else 502


@app.route('/upload', methods=['POST'])
def upload():
    global _state

    if 'file' not in request.files:
        return jsonify({'success': False, 'errors': ['No file part in request.']}), 400

    f = request.files['file']
    if not f or not f.filename:
        return jsonify({'success': False, 'errors': ['No file selected.']}), 400
    if not f.filename.lower().endswith('.docx'):
        return jsonify({'success': False, 'errors': ['Only .docx files are supported.']}), 400

    tmp_fd, tmp_path = tempfile.mkstemp(suffix='.docx')
    try:
        os.close(tmp_fd)
        f.save(tmp_path)

        parse_result = parse_document(tmp_path)

        if parse_result.errors:
            return jsonify({
                'success':  False,
                'errors':   parse_result.errors,
                'warnings': parse_result.warnings,
            }), 422

        use_ai = request.args.get('ai', 'true').lower() != 'false'
        glean_cfg = _state['glean_config'] if use_ai else None

        results = generate_all(
            parse_result.requirements,
            parse_result.specifications,
            parse_result.procedure_steps,
            use_ai=use_ai,
            glean_config=glean_cfg,
        )

        matched  = sum(1 for v in results.values() if not v['unmatched'])
        ai_count = sum(1 for v in results.values() if v.get('ai_generated'))
        stats = {
            'requirements':    len(parse_result.requirements),
            'specifications':  len(parse_result.specifications),
            'procedure_steps': len(parse_result.procedure_steps),
            'total':           len(results),
            'matched':         matched,
            'unmatched':       len(results) - matched,
        }

        _state.update({
            'results':          results,
            'filename':         f.filename,
            'stats':            stats,
            'warnings':         parse_result.warnings,
            '_requirements':    parse_result.requirements,
            '_specifications':  parse_result.specifications,
            '_procedure_steps': parse_result.procedure_steps,
        })

        return jsonify({
            'success':  True,
            'filename': f.filename,
            'warnings': parse_result.warnings,
            'stats':    stats,
            'ai_count': ai_count,
            'results':  results,
        })

    except Exception as exc:
        traceback.print_exc()
        return jsonify({'success': False,
                        'errors': [f'Unexpected error: {exc}']}), 500
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


@app.route('/regenerate-single', methods=['POST'])
def regenerate_single():
    """Re-generate AC for one item, with an optional custom prompt."""
    data          = request.get_json(silent=True) or {}
    number        = (data.get('number') or '').strip()
    custom_prompt = (data.get('custom_prompt') or '').strip()

    if not number:
        return jsonify({'success': False, 'error': 'number is required.'}), 400
    if not _state.get('_requirements') and not _state.get('_specifications'):
        return jsonify({'success': False, 'error': 'No document loaded.'}), 400
    if not ai_available(_state['glean_config']):
        return jsonify({'success': False, 'error': 'Glean AI is not configured.'}), 400

    # Locate the item in requirements or specifications
    item = next((r for r in _state['_requirements']   if r.number == number), None)
    if item is None:
        item = next((s for s in _state['_specifications'] if s.number == number), None)
    if item is None:
        return jsonify({'success': False, 'error': f'Item {number} not found in loaded document.'}), 404

    try:
        result = generate_acceptance_criteria_for(
            item,
            _state['_procedure_steps'],
            glean_config=_state['glean_config'],
            custom_prompt=custom_prompt,
        )
        # Merge into cached results
        if number in _state['results']:
            _state['results'][number].update({
                'generated_ac': result['generated_ac'],
                'ai_generated': result['ai_generated'],
            })
        return jsonify({
            'success':      True,
            'generated_ac': result['generated_ac'],
            'ai_generated': result['ai_generated'],
        })
    except Exception as exc:
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(exc)}), 500


@app.route('/regenerate', methods=['POST'])
def regenerate():
    global _state

    if not _state.get('_requirements') and not _state.get('_specifications'):
        return jsonify({'success': False,
                        'error': 'No document loaded. Please upload a file first.'}), 400

    if not ai_available(_state['glean_config']):
        return jsonify({'success': False,
                        'error': 'Glean AI is not configured. '
                                 'Please enter the service URL and API key first.'}), 400

    try:
        results = generate_all(
            _state['_requirements'],
            _state['_specifications'],
            _state['_procedure_steps'],
            use_ai=True,
            glean_config=_state['glean_config'],
        )

        matched  = sum(1 for v in results.values() if not v['unmatched'])
        ai_count = sum(1 for v in results.values() if v.get('ai_generated'))
        _state['results'] = results
        _state['stats']['matched'] = matched

        return jsonify({
            'success':  True,
            'results':  results,
            'ai_count': ai_count,
            'stats':    _state['stats'],
        })

    except Exception as exc:
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(exc)}), 500


@app.route('/export/excel')
def export_excel():
    if not _state['results']:
        return jsonify({'error': 'No results to export.'}), 400
    data = export_to_excel(_state['results'])
    stem = os.path.splitext(_state['filename'])[0] if _state['filename'] else 'results'
    return send_file(
        io.BytesIO(data),
        download_name=f'{stem}_acceptance_criteria.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@app.route('/export/docx')
def export_docx_route():
    if not _state['results']:
        return jsonify({'error': 'No results to export.'}), 400
    data = export_to_docx(_state['results'])
    stem = os.path.splitext(_state['filename'])[0] if _state['filename'] else 'results'
    return send_file(
        io.BytesIO(data),
        download_name=f'{stem}_acceptance_criteria.docx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )


# ---------------------------------------------------------------------------
# TC Review
# ---------------------------------------------------------------------------

@app.route('/review-tc', methods=['POST'])
def review_tc_route():
    """Upload a TC .docx and return a structured compliance review."""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file part in request.'}), 400

    f = request.files['file']
    if not f or not f.filename:
        return jsonify({'success': False, 'error': 'No file selected.'}), 400
    if not f.filename.lower().endswith('.docx'):
        return jsonify({'success': False, 'error': 'Only .docx files are supported.'}), 400

    tmp_fd, tmp_path = tempfile.mkstemp(suffix='.docx')
    try:
        os.close(tmp_fd)
        f.save(tmp_path)
        extra = _state.get('setup_config', {}).get('review_extra_guidance', '')
        result = reviewer.review_tc(tmp_path, _state['glean_config'], extra_prompt=extra)
        result['filename'] = f.filename
        return jsonify({'success': True, 'result': result})
    except Exception as exc:
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(exc)}), 500
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Setup config
# ---------------------------------------------------------------------------

@app.route('/api/setup-config', methods=['GET'])
def api_get_setup_config():
    """Return current setup config plus the built-in default AC prompt."""
    cfg = _state.get('setup_config', {'ac_system_prompt': '', 'review_extra_guidance': ''})
    return jsonify({
        'ac_system_prompt':      cfg.get('ac_system_prompt', ''),
        'review_extra_guidance': cfg.get('review_extra_guidance', ''),
        'default_ac_prompt':     DEFAULT_AC_SYSTEM_PROMPT,
    })


@app.route('/api/save-setup', methods=['POST'])
def api_save_setup():
    """Persist AC prompt + review guidance; apply prompt to summarizer."""
    data = request.get_json(silent=True) or {}
    cfg = {
        'ac_system_prompt':      (data.get('ac_system_prompt')      or '').strip(),
        'review_extra_guidance': (data.get('review_extra_guidance') or '').strip(),
    }
    _state['setup_config'] = cfg
    _save_setup_config(cfg)
    set_ac_system_prompt(cfg['ac_system_prompt'])
    return jsonify({'success': True, 'message': 'Setup configuration saved.'})


# ---------------------------------------------------------------------------
# Reference documents
# ---------------------------------------------------------------------------

@app.route('/api/references', methods=['GET'])
def api_references():
    """List files in the references/ folder."""
    try:
        files = [f for f in os.listdir(_REFERENCES_DIR) if not f.startswith('.')]
        files.sort()
        return jsonify({'files': files})
    except Exception as exc:
        return jsonify({'files': [], 'error': str(exc)})


@app.route('/api/upload-reference', methods=['POST'])
def api_upload_reference():
    """Upload a new document to the references/ folder."""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file part.'}), 400
    f = request.files['file']
    if not f or not f.filename:
        return jsonify({'success': False, 'error': 'No file selected.'}), 400

    filename = os.path.basename(f.filename)  # prevent path traversal
    try:
        os.makedirs(_REFERENCES_DIR, exist_ok=True)
        f.save(os.path.join(_REFERENCES_DIR, filename))
        return jsonify({'success': True, 'filename': filename})
    except Exception as exc:
        return jsonify({'success': False, 'error': str(exc)}), 500


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print()
    print("=" * 55)
    print("  SW Test Case Acceptance Criteria Builder")
    print(f"  Open your browser at:  http://localhost:{port}")
    print("=" * 55)
    print()
    app.run(debug=False, host='127.0.0.1', port=port)
