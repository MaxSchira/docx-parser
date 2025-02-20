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
    
    paragraph_counter = 0

    for para in doc.paragraphs:
        paragraph_counter += 1  
        paragraph_data = {
            "paragraph_id": paragraph_counter,
            "is_empty": len(para.text.strip()) == 0,
            "runs": []
        }

        for run in para.runs:
            text = run.text.strip()
            if not text:
                continue  

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
import logging

@app.route('/generate-docx', methods=['POST'])
def generate_docx():
    try:
        # JSON-String aus form-data abrufen, falls nicht direkt als JSON gesendet
        raw_data = request.form.get("updated_speisekarte")
        
        if raw_data:
            app.logger.info("Received raw_data (string): %s", raw_data)  # Logge den rohen Input
            updated_speisekarte = json.loads(raw_data)  # String in JSON umwandeln
        else:
            app.logger.error("Fehlende Speisekarte-Daten")
            return jsonify({"error": "Fehlende Speisekarte-Daten"}), 400

        # Debug-Log: Struktur der empfangenen Speisekarte
        app.logger.info("Parsed updated_speisekarte JSON: %s", json.dumps(updated_speisekarte, indent=2))

        # Datei abrufen
        file = request.files.get("file")
        if not file:
            app.logger.error("Fehlende DOCX-Datei")
            return jsonify({"error": "Fehlende DOCX-Datei"}), 400

        doc = Document(file)

        # Entferne alle bestehenden Absätze
        while len(doc.paragraphs) > 0:
            p = doc.paragraphs[0]
            p._element.getparent().remove(p._element)

        # Aktualisierte Speisekarte einfügen
        for item in updated_speisekarte:
            if item.get("is_empty", False):
                doc.add_paragraph("")
                continue
            
            para = doc.add_paragraph()
            for run_data in item.get("runs", []):
                run = para.add_run(run_data["text"])
                run.bold = run_data["style"]["bold"]
                run.italic = run_data["style"]["italic"]
                run.underline = run_data["style"]["underline"]
                run.font.name = run_data["style"]["font"]
                run.font.size = Pt(run_data["style"]["size"])

                # Farbe setzen
                if run_data["style"]["color"] != "#000000":
                    rgb = run_data["style"]["color"].lstrip("#")
                    run.font.color.rgb = RGBColor(int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16))

        # Datei in Speicher statt auf Festplatte speichern
        output = BytesIO()
        doc.save(output)
        output.seek(0)

        return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                         as_attachment=True, download_name="updated_menu.docx")

    except Exception as e:
        app.logger.error("Fehler: %s", str(e))
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
