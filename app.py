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
    data = request.get_json()
    updated_speisekarte = data.get("updated_speisekarte", [])

    doc = Document()
    
    for item in updated_speisekarte:
        para = doc.add_paragraph()
        run = para.add_run(item["text"])
        
        # **Formatierung aus JSON übernehmen**
        style = item["style"]
        run.bold = style["bold"]
        run.italic = style["italic"]
        run.underline = style["underline"]
        run.font.name = style["font"]
        run.font.size = Pt(style["size"])
        if style["color"] and style["color"] != "#000000":
            rgb_color = RGBColor(int(style["color"][1:3], 16), int(style["color"][3:5], 16), int(style["color"][5:7], 16))
            run.font.color.rgb = rgb_color

    # **Datei in den Speicher speichern**
    output = BytesIO()
    doc.save(output)
    output.seek(0)

    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     as_attachment=True, download_name="updated_menu.docx")

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
