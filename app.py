
from flask import Flask, request, send_file, make_response
from flask_cors import CORS
from docx import Document
from docx.shared import Pt
import io

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

    def replace_last_bullets(section_title, new_bullets, count):
        for i, para in enumerate(doc.paragraphs):
            if section_title in para.text:
                bullet_indices = []
                j = i + 1
                while j < len(doc.paragraphs):
                    if doc.paragraphs[j].text.strip().startswith("•"):
                        bullet_indices.append(j)
                    elif doc.paragraphs[j].text.strip() == "" or doc.paragraphs[j].text.strip()[0].isupper():
                        break
                    j += 1
                for k in range(count):
                    idx = bullet_indices[-count + k]
                    doc.paragraphs[idx].text = new_bullets[k]
                    for run in doc.paragraphs[idx].runs:
                        run.font.size = Pt(10.5)
                        run.font.name = "Times New Roman"
                break

    replace_last_bullets("iCONSULT COLLABORATIVE", experience[:1], 1)
    replace_last_bullets("FRAPPE TECHNOLOGIES PRIVATE LIMITED", experience[1:3], 2)
    replace_last_bullets("ERNST & YOUNG", experience[3:4], 1)

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
    app.run(debug=True)
