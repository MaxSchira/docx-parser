from flask import Flask, request, jsonify
from docx import Document

app = Flask(__name__)

@app.route('/parse-docx', methods=['POST'])
def parse_docx():
    file = request.files['file']
    doc = Document(file)
    extracted_text = []

    # 1️⃣ Fließtext extrahieren (inkl. Formatierungen)
    for para in doc.paragraphs:
        text = para.text.strip()
        styles = []
        for run in para.runs:
            if run.bold:
                styles.append("bold")
            if run.italic:
                styles.append("italic")
            if run.underline:
                styles.append("underline")
        if text:
            extracted_text.append({"text": text, "style": styles if styles else ["paragraph"]})

    # 2️⃣ Tabellen auslesen (inkl. Formatierung)
    for table in doc.tables:
        for row in table.rows:
            row_text = []
            row_styles = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                cell_styles = []
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.bold:
                            cell_styles.append("bold")
                        if run.italic:
                            cell_styles.append("italic")
                        if run.underline:
                            cell_styles.append("underline")
                if cell_text:
                    row_text.append(cell_text)
                    row_styles.append(cell_styles if cell_styles else ["table"])
            if row_text:
                extracted_text.append({"text": " | ".join(row_text), "style": row_styles})

    # 3️⃣ Text aus Textboxen extrahieren (inkl. Formatierung)
    for shape in doc.element.xpath("//w:p"):
        text = " ".join([t.text for t in shape.xpath(".//w:t") if t.text]).strip()
        styles = []
        for run in shape.xpath(".//w:r"):
            if run.xpath(".//w:b"):
                styles.append("bold")
            if run.xpath(".//w:i"):
                styles.append("italic")
            if run.xpath(".//w:u"):
                styles.append("underline")
        if text:
            extracted_text.append({"text": text, "style": styles if styles else ["textbox"]})

    return jsonify(extracted_text)

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
