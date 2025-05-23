
from flask import Flask, request, send_file, make_response
from flask_cors import CORS
from docx import Document
from docx.shared import Pt
import io
import os
from datetime import datetime

app = Flask(__name__)
CORS(app, resources={r"/tailor": {"origins": "chrome-extension://*"}})

@app.route('/tailor', methods=['POST', 'OPTIONS'])
def tailor_resume():
    origin = request.headers.get("Origin", "*")

    if request.method == "OPTIONS":
        response = make_response()
        response.headers["Access-Control-Allow-Origin"] = origin
        response.headers["Access-Control-Allow-Headers"] = "Content-Type"
        response.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        return response

    data = request.json
    experience = data.get("experience", [])
    skills = data.get("skills", "")

    doc = Document("base_resume.docx")

    def replace_bullets_by_match(section_title, new_bullets, count):
        found_index = None
        for i, para in enumerate(doc.paragraphs):
            if section_title in para.text:
                found_index = i
                break

        if found_index is None:
            print("❌ Section '{}' not found.".format(section_title))
            return

        bullet_indices = []
        for j in range(found_index + 1, len(doc.paragraphs)):
            text = doc.paragraphs[j].text.strip()
            if len(text) > 0 and text.isupper():
                break
            if text.startswith("•"):
                bullet_indices.append(j)

        if len(bullet_indices) < count:
            print("⚠️ Not enough bullet lines under '{}'. Found {}.".format(section_title, len(bullet_indices)))
            return

        for k in range(count):
            idx = bullet_indices[k]
            doc.paragraphs[idx].text = new_bullets[k]
            for run in doc.paragraphs[idx].runs:
                run.font.size = Pt(10.5)
                run.font.name = "Times New Roman"

    replace_bullets_by_match("iCONSULT COLLABORATIVE", experience[:1], 1)
    replace_bullets_by_match("FRAPPE TECHNOLOGIES PRIVATE LIMITED", experience[1:3], 2)
    replace_bullets_by_match("ERNST & YOUNG", experience[3:4], 1)

    for i, para in enumerate(doc.paragraphs):
        if "Core Competencies" in para.text:
            para.text += " | " + skills
            break

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    filename = "Rijul Chaturvedi Resume - " + datetime.now().strftime('%Y-%m-%d') + ".docx"

    response = make_response(send_file(
        output,
        download_name=filename,
        as_attachment=True
    ))
    response.headers["Access-Control-Allow-Origin"] = origin
    return response

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)

