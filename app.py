
from flask import Flask, request, send_file, make_response
from flask_cors import CORS
from docx import Document
from docx.shared import Pt
import io
import os

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

    def replace_last_n_after_heading(section_title, new_texts, count):
        if len(new_texts) < count:
            print(f"⚠️ Not enough GPT bullets provided for '{section_title}' — expected {count}, got {len(new_texts)}")
            return

        found_index = None
        for i, para in enumerate(doc.paragraphs):
            if section_title in para.text:
                found_index = i
                break

        if found_index is None:
            print(f"❌ Section '{section_title}' not found.")
            return

        section_paras = []
        for j in range(found_index + 1, len(doc.paragraphs)):
            text = doc.paragraphs[j].text.strip()
            if len(text) > 0 and text.isupper():  # likely a new section heading
                break
            if text:
                section_paras.append(j)

        if len(section_paras) < count:
            print(f"⚠️ Not enough resume paragraphs under '{section_title}' to replace {count}.")
            return

        for k in range(count):
            idx = section_paras[-count + k]
            doc.paragraphs[idx].text = new_texts[k]
            for run in doc.paragraphs[idx].runs:
                run.font.size = Pt(10.5)
                run.font.name = "Times New Roman"

    replace_last_n_after_heading("iCONSULT COLLABORATIVE", experience[:1], 1)
    replace_last_n_after_heading("FRAPPE TECHNOLOGIES PRIVATE LIMITED", experience[1:3], 2)
    replace_last_n_after_heading("ERNST & YOUNG", experience[3:4], 1)

    for i, para in enumerate(doc.paragraphs):
        if "Core Competencies" in para.text:
            para.text += " | " + skills
            break

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    response = make_response(send_file(
        output,
        download_name="Rijul_Chaturvedi_Resume_Tailored.docx",
        as_attachment=True
    ))
    response.headers["Access-Control-Allow-Origin"] = origin
    return response

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)
