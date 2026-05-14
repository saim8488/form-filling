from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.oxml.ns import qn
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
        template = request.files.get('template')
        if not template:
            return jsonify({"error": "Missing template file"}), 400

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

        doc = Document(template)
        pattern = re.compile(r'\{\{\s*(\w+)\s*\}\}')

        def do_replace(match):
            return str(data.get(match.group(1), match.group(0)))

        def process_all_paragraphs(root_element):
            """Walk every <w:p> anywhere in the tree, including inside textboxes/shapes."""
            for p in root_element.iter(qn('w:p')):
                t_elements = p.findall('.//' + qn('w:t'))
                if not t_elements:
                    continue
                full_text = ''.join(t.text or '' for t in t_elements)
                if '{{' not in full_text:
                    continue
                new_text = pattern.sub(do_replace, full_text)
                if new_text != full_text:
                    t_elements[0].text = new_text
                    for t in t_elements[1:]:
                        t.text = ''

        # Body — catches paragraphs, tables, AND textboxes anywhere
        process_all_paragraphs(doc.element.body)

        # Headers and footers
        for section in doc.sections:
            process_all_paragraphs(section.header._element)
            process_all_paragraphs(section.footer._element)

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
