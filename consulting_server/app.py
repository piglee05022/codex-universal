from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from tempfile import NamedTemporaryFile
from legal_module.filing import create_filing
import json
import os

app = Flask(__name__)
app.secret_key = "secret-key"

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            case_number = request.form.get('case_number', '')
            parties = request.form.get('parties', '')
            court = request.form.get('court', '')
            claims = request.form.get('claims', '')
            facts = request.form.get('facts', '')
            laws = [l.strip() for l in request.form.get('laws', '').splitlines() if l.strip()]
            evidence_json = request.form.get('evidence', '[]')
            attachments_json = request.form.get('attachments', '[]')
            evidence = json.loads(evidence_json) if evidence_json else []
            attachments = json.loads(attachments_json) if attachments_json else []
            case_info = {
                'case_number': case_number,
                'parties': parties,
                'court': court,
                'claims': claims,
                'facts': facts,
                'laws': laws,
                'evidence': evidence,
                'attachments': attachments,
            }
            with NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
                tmp_path = tmp.name
            create_filing(case_info, tmp_path)
            return send_file(tmp_path, as_attachment=True, download_name='filing.docx')
        except Exception as exc:
            flash(f'Error generating document: {exc}')
            return redirect(url_for('index'))

    return render_template('form.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
