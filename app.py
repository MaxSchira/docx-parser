from flask import Flask, request, jsonify
from docx import Document

app = Flask(__name__)

@app.route('/parse-docx', methods=['POST'])
def parse_docx():
    file = request.files['file']
    doc = Document(file)
    extracted_text = []

    # 1️ Fließtext extrahieren
    for para in doc.paragraphs:
        extracted_text.append({"text": para.text, "style": []})

    # 2️ Text aus Textboxen extrahieren
    for shape in doc.inline_shapes:
        if shape._element.xpath(".//w:t"):
            text = " ".join([t.text for t in shape._element.xpath(".//w:t")])
            extracted_text.append({"text": text, "style": ["textbox"]})

    # 3 Text aus Tabellen extrahieren
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                extracted_text.append({"text": cell.text, "style": ["table"]})

    return jsonify(extracted_text)

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
