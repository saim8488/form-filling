from flask import Flask, request, send_file, jsonify
from docx import Document
from io import BytesIO
import re
import json
import os
import logging

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

API_KEY = os.environ.get('API_KEY')


@app.before_request
def check_auth():
    if request.path in ('/', '/health'):
        return
    if API_KEY and request.headers.get('X-API-Key') != API_KEY:
        return jsonify({"error": "Unauthorized"}), 401


@app.get("/")
def health():
    return "ok"


@app.get("/health")
def health_alt():
    return jsonify({"status": "ok"})


@app.post("/fill-docx")
def fill_docx():
    try:
        # Get the template file
        template = request.files.get('template')
        if not template:
            return jsonify({"error": "Missing template file"}), 400

        # Parse the data field — it arrives as a JSON STRING in form-data
        data = {}
        data_raw = request.form.get('data')
        if data_raw:
            try:
                data = json.loads(data_raw)
            except json.JSONDecodeError as e:
                return jsonify({"error": f"Invalid JSON in data field: {e}"}), 400
        elif request.is_json:
            data = request.get_json()
        elif 'data' in request.files:
            data = json.loads(request.files['data'].read().decode('utf-8'))

        if not isinstance(data, dict):
            return jsonify({"error": "data must be a JSON object"}), 400

        app.logger.info(f"Filling template with {len(data)} fields")
        app.logger.info(f"Sample keys: {list(data.keys())[:5]}")

        # Open the DOCX
        doc = Document(template)

        def replace_in_paragraph(p):
            """Replace all {{placeholders}} in a paragraph, handling split runs."""
            if not p.runs:
                return
            full = ''.join(r.text for r in p.runs)
            if '{{' not in full:
                return  # quick skip
            new = re.sub(
                r'\{\{\s*(\w+)\s*\}\}',
                lambda m: str(data.get(m.group(1), m.group(0))),
                full
            )
            if new != full:
                # Clear all runs except the first
                for r in p.runs[1:]:
                    r.text = ''
                p.runs[0].text = new

        # Process all paragraphs in the document body
        for p in doc.paragraphs:
            replace_in_paragraph(p)

        # Process all tables (including nested)
        def process_tables(tables):
            for tbl in tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            replace_in_paragraph(p)
                        # Recurse into nested tables
                        if cell.tables:
                            process_tables(cell.tables)

        process_tables(doc.tables)

        # Process headers and footers
        for section in doc.sections:
            for hdr_ftr in (section.header, section.footer):
                for p in hdr_ftr.paragraphs:
                    replace_in_paragraph(p)
                for tbl in hdr_ftr.tables:
                    for row in tbl.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                replace_in_paragraph(p)

        # Save and return
        out = BytesIO()
        doc.save(out)
        out.seek(0)

        filename = data.get('_filename', 'filled.docx')
        return send_file(
            out,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        app.logger.error(f"Fill error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
