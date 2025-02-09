from flask import Flask, request, jsonify
from docx import Document
from docx.shared import RGBColor

app = Flask(__name__)

@app.route('/parse-docx', methods=['POST'])
def parse_docx():
    file = request.files['file']
    doc = Document(file)
    extracted_text = []

    # Debugging-Zähler
    debug_counts = {
        "paragraphs_extracted": 0,
        "tables_extracted": 0,
        "textboxes_extracted": 0
    }

    # 1️⃣ Fließtext & Formatierungen extrahieren
    for para in doc.paragraphs:
        text = para.text.strip()
        styles = {
            "font": None,
            "size": None,
            "bold": False,
            "italic": False,
            "underline": False,
            "alignment": str(para.alignment) if para.alignment else "left",
            "color": "#000000"
        }
        for run in para.runs:
            if run.font:
                styles["font"] = run.font.name if run.font.name else "Default"
                styles["size"] = run.font.size.pt if run.font.size else None
                styles["bold"] = run.bold if run.bold else False
                styles["italic"] = run.italic if run.italic else False
                styles["underline"] = run.underline if run.underline else False
                if run.font.color and isinstance(run.font.color.rgb, RGBColor):
                    styles["color"] = f"#{run.font.color.rgb}"  # RGB zu Hex umwandeln

        if text:
            extracted_text.append({"text": text, "type": "paragraph", "style": styles})
            debug_counts["paragraphs_extracted"] += 1  # Debug-Zähler erhöhen

    # 2️⃣ Tabelleninhalte extrahieren (inkl. Formatierung)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_text = cell.text.strip()
                cell_styles = {
                    "font": None,
                    "size": None,
                    "bold": False,
                    "italic": False,
                    "underline": False,
                    "alignment": "left",
                    "color": "#000000"
                }
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.font:
                            cell_styles["font"] = run.font.name if run.font.name else "Default"
                            cell_styles["size"] = run.font.size.pt if run.font.size else None
                            cell_styles["bold"] = run.bold if run.bold else False
                            cell_styles["italic"] = run.italic if run.italic else False
                            cell_styles["underline"] = run.underline if run.underline else False
                            if run.font.color and isinstance(run.font.color.rgb, RGBColor):
                                cell_styles["color"] = f"#{run.font.color.rgb}"

                if cell_text:
                    extracted_text.append({"text": cell_text, "type": "table", "style": cell_styles})
                    debug_counts["tables_extracted"] += 1  # Debug-Zähler erhöhen

    # 3️⃣ Text aus Textboxen extrahieren (inkl. Formatierung)
    for shape in doc.element.xpath("//w:p"):
        text = " ".join([t.text for t in shape.xpath(".//w:t") if t.text]).strip()
        styles = {
            "font": None,
            "size": None,
            "bold": False,
            "italic": False,
            "underline": False,
            "alignment": "left",
            "color": "#000000"
        }
        for run in shape.xpath(".//w:r"):
            if run.xpath(".//w:b"):
                styles["bold"] = True
            if run.xpath(".//w:i"):
                styles["italic"] = True
            if run.xpath(".//w:u"):
                styles["underline"] = True

        if text:
            extracted_text.append({"text": text, "type": "textbox", "style": styles})
            debug_counts["textboxes_extracted"] += 1  # Debug-Zähler erhöhen

    # Debugging-Output zurückgeben
    return jsonify({
        "extracted_data": extracted_text,
        "debug_info": debug_counts
    })

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
