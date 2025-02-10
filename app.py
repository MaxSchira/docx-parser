from flask import Flask, request, jsonify
from docx import Document
from docx.shared import RGBColor

app = Flask(__name__)

@app.route('/parse-docx', methods=['POST'])
def parse_docx():
    file = request.files['file']
    doc = Document(file)
    extracted_text = []

    for para in doc.paragraphs:
        for run in para.runs:
            text = run.text.strip()
            if not text:
                continue  # Leere Runs ignorieren

            styles = {
                "font": run.font.name if run.font and run.font.name else "Calibri",
                "size": run.font.size.pt if run.font and run.font.size else None,
                "bold": run.bold if run.bold else False,
                "italic": run.italic if run.italic else False,
                "underline": run.underline if run.underline else False,
                "color": f"#{run.font.color.rgb}" if run.font.color and isinstance(run.font.color.rgb, RGBColor) else "#000000"
            }

            extracted_text.append({"text": text, "type": "paragraph", "style": styles})

    return jsonify({"extracted_data": extracted_text})

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
