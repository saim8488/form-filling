from flask import Flask, request, send_file
from docx import Document
from io import BytesIO
import re

app = Flask(__name__)

@app.post("/fill-docx")
def fill_docx():
    template = request.files['template']
    data = request.json if request.is_json else request.form.to_dict()
    if 'data' in request.files:
        import json
        data = json.loads(request.files['data'].read())
    
    doc = Document(template)
    
    def replace_in_paragraph(p):
        full = ''.join(r.text for r in p.runs)
        new = re.sub(r'\{\{(\w+)\}\}', lambda m: str(data.get(m.group(1), m.group(0))), full)
        if new != full and p.runs:
            for r in p.runs[1:]:
                r.text = ''
            p.runs[0].text = new
    
    for p in doc.paragraphs:
        replace_in_paragraph(p)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p)
    
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return send_file(out, download_name='filled.docx',
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.get("/")
def health():
    return "ok"
