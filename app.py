import json

from flask import Flask, request, jsonify, send_file
from docx import Document
from docx.shared import RGBColor, Pt
from io import BytesIO

app = Flask(__name__)

# ENDPOINT 1: DOCX-PARSING
@app.route('/parse-docx', methods=['POST'])
def parse_docx():
    file = request.files['file']
    doc = Document(file)
    extracted_text = []
    
    # Absatz-Zähler für ID-Zuweisung
    paragraph_counter = 0

    for para in doc.paragraphs:
        paragraph_counter += 1  # ID für jeden Absatz
        paragraph_data = {
            "paragraph_id": paragraph_counter,
            "is_empty": len(para.text.strip()) == 0,  # Prüfen, ob Absatz leer ist
            "runs": []  # Enthält alle Runs dieses Absatzes
        }

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

            paragraph_data["runs"].append({"text": text, "style": styles})

        extracted_text.append(paragraph_data)

    return jsonify({"extracted_data": extracted_text})

# ENDPOINT 2: TEXT ERSETZEN UND DOCX GENERIEREN
@app.route('/generate-docx', methods=['POST'])
def generate_docx():
    try:
        # JSON sicher parsen, falls es als String kommt
        raw_data = request.get_json()
        if isinstance(raw_data, str):
            raw_data = json.loads(raw_data)

        updated_speisekarte = raw_data.get("updated_speisekarte", [])

        # Bestehendes Dokument öffnen
        file = request.files["file"]
        doc = Document(file)

        # Bestehenden Text leeren
        for para in doc.paragraphs:
            para.clear()

        # Aktualisierte Speisekarte einfügen
        doc.paragraphs.clear()  # Vorherige Absätze entfernen
        for item in updated_speisekarte:
            para = doc.add_paragraph()
            
            if item.get("is_empty", False):  
                para.add_run("")  # Leere Zeile
                continue
            
            for run_data in item.get("runs", []):
                run = para.add_run(run_data["text"])
                run.bold = run_data["style"]["bold"]
                run.italic = run_data["style"]["italic"]
                run.underline = run_data["style"]["underline"]
                run.font.name = run_data["style"]["font"]
                run.font.size = Pt(run_data["style"]["size"])
                
                if run_data["style"]["color"] != "#000000":
                    rgb = run_data["style"]["color"].lstrip("#")
                    run.font.color.rgb = RGBColor(int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16))

        # Datei in den Speicher statt auf die Festplatte speichern
        output = BytesIO()
        doc.save(output)
        output.seek(0)

        return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                         as_attachment=True, download_name="updated_menu.docx")

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
