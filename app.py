from flask import Flask, request, jsonify
from docx import Document

app = Flask(__name__)

@app.route('/parse-docx', methods=['POST'])
def parse_docx():
    file = request.files['file']
    doc = Document(file)
    extracted_text = []

    for para in doc.paragraphs:
        styles = []
        if para.runs:
            for run in para.runs:
                if run.bold:
                    styles.append("bold")
                if run.italic:
                    styles.append("italic")
                if run.underline:
                    styles.append("underline")
        extracted_text.append({"text": para.text, "style": styles})

    return jsonify(extracted_text)

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)