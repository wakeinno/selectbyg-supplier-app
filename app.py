"""
Select by G Group — Supplier Presentation Generator
Flask web application for collecting supplier information and generating PPT presentations.
"""

import os
import sys
import shutil
import uuid
from pathlib import Path
from flask import Flask, render_template, request, jsonify

sys.path.insert(0, os.path.dirname(__file__))
from generate_pptx import generate_presentation

app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = 'selectbyg-secret-2024'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB max

UPLOAD_FOLDER = Path(__file__).parent / 'uploads'
UPLOAD_FOLDER.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    session_id = str(uuid.uuid4())
    session_dir = UPLOAD_FOLDER / session_id
    session_dir.mkdir(exist_ok=True)

    try:
        data = {
            'company_name': request.form.get('company_name', '').strip(),
            'category':     request.form.get('category', '').strip(),
            'history':      request.form.get('history', '').strip(),
            'identity':     request.form.get('identity', '').strip(),
            'projects':     request.form.get('projects', '').strip(),
            'added_values': [],
            'references':   [],
        }

        if not data['company_name']:
            return jsonify({'error': 'Company name is required'}), 400

        for i in range(1, 21):
            val = request.form.get(f'added_value_{i}', '').strip()
            if val:
                data['added_values'].append(val)

        for i in range(1, 9):
            ref_text = request.form.get(f'reference_{i}', '').strip()
            if not ref_text:
                continue
            stars_raw = request.form.get(f'reference_stars_{i}', '0')
            try:
                star_count = max(0, min(5, int(stars_raw)))
            except (ValueError, TypeError):
                star_count = 0
            stars_str = '\u2605' * star_count
            combined = f"{ref_text} {stars_str}".strip() if stars_str else ref_text
            data['references'].append(combined)

        supplier_logo_path = None
        logo_file = request.files.get('supplier_logo')
        if logo_file and logo_file.filename and allowed_file(logo_file.filename):
            ext = logo_file.filename.rsplit('.', 1)[1].lower()
            logo_dest = session_dir / f"supplier_logo.{ext}"
            logo_file.save(str(logo_dest))
            supplier_logo_path = str(logo_dest)

        def save_file(file, name):
            if file and file.filename and allowed_file(file.filename):
                ext = file.filename.rsplit('.', 1)[1].lower()
                dest = session_dir / f"{name}.{ext}"
                file.save(str(dest))
                return str(dest)
            return None

        supplier_photos = []
        for i in [1, 2]:
            path = save_file(request.files.get(f'supplier_photo_{i}'), f'supplier_{i}')
            if path:
                supplier_photos.append(path)

        photo_groups = []
        hotel_count = int(request.form.get('hotel_count', 0))
        for i in range(1, hotel_count + 1):
            hotel_name = request.form.get(f'hotel_photo_name_{i}', '').strip()
            path = save_file(request.files.get(f'hotel_photo_{i}'), f'hotel_{i}')
            if hotel_name and path:
                photo_groups.append({
                    'hotel_name': hotel_name,
                    'photos':     [path],
                })

        output_path = generate_presentation(
            data, photo_groups,
            supplier_logo_path=supplier_logo_path,
            supplier_photos=supplier_photos,
        )

        return jsonify({
            'success':  True,
            'message':  'Presentation generated successfully!',
            'filename': output_path.name,
            'path':     str(output_path),
        })

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500

    finally:
        try:
            shutil.rmtree(session_dir)
        except Exception:
            pass


@app.errorhandler(413)
def request_entity_too_large(e):
    return jsonify({'error': 'Files too large. Maximum upload size is 50 MB total.'}), 413


@app.errorhandler(500)
def internal_server_error(e):
    return jsonify({'error': 'Internal server error. Check the terminal for details.'}), 500


@app.route('/health')
def health():
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    print(f"\n\U0001f680 Select by G Presentation Generator running on http://localhost:{port}\n")
    app.run(debug=False, port=port, host='0.0.0.0')
