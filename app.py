import json

from flask import Flask, request, jsonify, send_file
from docx import Document
from docx.shared import RGBColor, Pt
from io import BytesIO

app = Flask(__name__)

#  ENDPOINT 1: DOCX-PARSING
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
                "size": run.font.size.pt if run.font and run.font.size else 12,
                "bold": run.bold if run.bold else False,
                "italic": run.italic if run.italic else False,
                "underline": run.underline if run.underline else False,
                "color": f"#{run.font.color.rgb}" if run.font.color and isinstance(run.font.color.rgb, RGBColor) else "#000000"
            }

            extracted_text.append({"text": text, "type": "paragraph", "style": styles})

    return jsonify({"extracted_data": extracted_text})

#  ENDPOINT 2: TEXT ERSETZEN UND DOCX GENERIEREN
@app.route('/generate-docx', methods=['POST'])
def generate_docx():
    try:
        # JSON sicher parsen, falls es als String kommt
        raw_data = request.get_json()
        if isinstance(raw_data, str):
            raw_data = json.loads(raw_data)

        updated_speisekarte = raw_data.get("updated_speisekarte", [])

        # Statt ein neues Dokument zu erstellen -> Bestehendes öffnen
        file = request.files["file"]
        doc = Document(file)

        # Bestehenden Text entfernen
        for para in doc.paragraphs:
            for run in para.runs:
                run.text = ""

        # Aktualisierten Text in bestehende Paragraphs einfügen
        for para, item in zip(doc.paragraphs, updated_speisekarte):
            run = para.add_run(item["text"])
            run.bold = item["style"]["bold"]
            run.italic = item["style"]["italic"]
            run.underline = item["style"]["underline"]
            run.font.name = item["style"]["font"]
            run.font.size = Pt(item["style"]["size"])
            if item["style"]["color"] != "#000000":
                rgb = item["style"]["color"].lstrip("#")
                run.font.color.rgb = RGBColor(int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16))

        # **Datei in den Speicher statt auf die Festplatte**
        output = BytesIO()
        doc.save(output)
        output.seek(0)

        return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                         as_attachment=True, download_name="updated_menu.docx")

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
