from flask import Flask, request, jsonify
from docx import Document

app = Flask(__name__)

@app.route('/parse-docx', methods=['POST'])
def parse_docx():
    file = request.files['file']
    doc = Document(file)
    extracted_text = []

    # 1️⃣ Fließtext extrahieren (Menü, Titel)
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:  # Leere Zeilen ignorieren
            extracted_text.append({"text": text, "style": ["paragraph"]})

    # 2️⃣ Text aus Tabellen extrahieren (Getränke, Weine)
    for table in doc.tables:
        for row in table.rows:
            row_text = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                if cell_text:
                    row_text.append(cell_text)
            if row_text:
                extracted_text.append({"text": " | ".join(row_text), "style": ["table"]})

    return jsonify(extracted_text)

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
