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
        # Debugging-Logs
        print("\nREQUEST RECEIVED")
        print("Received Headers:", request.headers)
        print("Received Content-Type:", request.content_type)
        print("Received Form Keys:", request.form.keys())

        # JSON sicher parsen, falls es als String kommt
        if request.content_type == "application/json":
            raw_data = request.get_json()
        elif "multipart/form-data" in request.content_type:
            raw_data = json.loads(request.form["updated_speisekarte"])
        else:
            return jsonify({"error": "Unsupported Media Type"}), 415

        # Falls raw_data eine Liste ist, dann setze sie direkt als updated_speisekarte
        if isinstance(raw_data, list):
            updated_speisekarte = raw_data
        else:
            updated_speisekarte = raw_data.get("updated_speisekarte", [])

        # Bestehendes Dokument öffnen
        file = request.files.get("file")
        if not file:
            return jsonify({"error": "No file provided!"}), 400
        
        doc = Document(file)

        # Bestehenden Text entfernen
        for para in doc.paragraphs:
            for run in para.runs:
                run.text = ""
                
        
        # Alle leeren Absätze sammeln
        empty_paragraphs = [p for p in doc.paragraphs if not any(run.text.strip() for run in p.runs)]

        # Leere Absätze aus dem Dokument entfernen
        for p in empty_paragraphs:
            p._element.getparent().remove(p._element)

        # Aktualisierten Text in bestehende Paragraphs einfügen
        current_para = None  # Aktueller Absatz speichern

        for item in updated_speisekarte:
            # Falls es ein Paragraph ist oder der Absatz nicht existiert, neuen Absatz beginnen
            if current_para is None or item["type"] == "paragraph":
                current_para = doc.add_paragraph()
    
            run = current_para.add_run(item["text"])
            run.bold = item["style"]["bold"]
            run.italic = item["style"]["italic"]
            run.underline = item["style"]["underline"]
            run.font.name = item["style"]["font"]
            run.font.size = Pt(item["style"]["size"])

            # Falls eine spezielle Farbe gesetzt ist, übernehmen
            if item["style"]["color"] != "#000000":
                rgb = item["style"]["color"].lstrip("#")
                run.font.color.rgb = RGBColor(int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16))

        # Datei im Speicher statt auf der Festplatte speichern
        output = BytesIO()
        doc.save(output)
        output.seek(0)

        return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                         as_attachment=True, download_name="updated_menu.docx")

    except Exception as e:
        print("\nERROR OCCURRED")
        print(str(e))
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
