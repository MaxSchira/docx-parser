from flask import Flask, request, jsonify
from docx import Document
from lxml import etree

app = Flask(__name__)

@app.route('/parse-docx', methods=['POST'])
def parse_docx():
    file = request.files['file']
    doc = Document(file)
    extracted_text = []

    # 1️⃣ Fließtext extrahieren (normale Absätze)
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            extracted_text.append({"text": text, "style": ["paragraph"]})

    # 2️⃣ Tabellen auslesen (z. B. Getränke- und Weinliste)
    for table in doc.tables:
        for row in table.rows:
            row_text = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                if cell_text:
                    row_text.append(cell_text)
            if row_text:
                extracted_text.append({"text": " | ".join(row_text), "style": ["table"]})

    # 3️⃣ Text aus Textboxen extrahieren
    for shape in doc.element.xpath("//w:p"):  # Alle Absätze in Shapes durchsuchen
        text = " ".join([t.text for t in shape.xpath(".//w:t") if t.text]).strip()
        if text:
            extracted_text.append({"text": text, "style": ["textbox"]})

    return jsonify(extracted_text)

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
