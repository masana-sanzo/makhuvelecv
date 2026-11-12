import os
import secrets
from io import BytesIO

from flask import Flask, abort, redirect, render_template, request, send_file, send_from_directory, session, url_for
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or secrets.token_hex(32)

STEPS = {
    1: {
        "title": "Personal Information",
        "next_label": "Next Page: Education",
        "fields": [
            {"name": "name", "label": "Name", "value": "Makhuvele Masana"},
            {"name": "email", "label": "Email", "type": "email", "value": "sanzokhubs@gamil.com"},
            {
                "name": "bio",
                "label": "Bio",
                "type": "textarea",
                "value": "Experienced financial analyst with expertise in financial modeling, data analysis, and business strategy.",
            },
        ],
    },
    2: {
        "title": "Education",
        "next_label": "Next Page: Work Experience",
        "fields": [
            {"name": "degree", "label": "Degree", "value": "BSC In Computer science"},
            {"name": "institution", "label": "Institution", "value": "University of Johannesburg"},
            {"name": "year", "label": "Year of Graduation", "value": "2021"},
        ],
    },
    3: {
        "title": "Work Experience",
        "next_label": "Next Page: Skills",
        "fields": [
            {"name": "jobTitle", "label": "Job Title", "value": "Software Programmer"},
            {"name": "company", "label": "Company", "value": "Sanzo Media "},
            {"name": "duration", "label": "Duration", "value": "2022-present"},
        ],
    },
    4: {
        "title": "Skills",
        "next_label": "Next: Projects",
        "fields": [
            {
                "name": "skills",
                "label": "Skills (comma-separated)",
                "value": "database analyst, software Programmer, website creator",
            },
        ],
    },
    5: {
        "title": "Projects",
        "next_label": "Next: Contact",
        "fields": [
            {"name": "projects", "label": "Projects (comma-separated)", "value": "Project 1, Project 2, Project 3"},
        ],
    },
    6: {
        "title": "Contact",
        "next_label": "Show all info in 1 page",
        "fields": [
            {"name": "contactInfo", "label": "Contact Info", "value": "contact number: 082 457 4898"},
        ],
    },
}

SUMMARY_FIELDS = [
    ("Name", "name"),
    ("Email", "email"),
    ("Bio", "bio"),
    ("Degree", "degree"),
    ("Institution", "institution"),
    ("Year of Graduation", "year"),
    ("Job Title", "jobTitle"),
    ("Company", "company"),
    ("Duration", "duration"),
    ("Skills", "skills"),
    ("Projects", "projects"),
    ("Contact Info", "contactInfo"),
]


@app.get("/")
def index():
    return render_template("index.html")


@app.route("/step/<int:step_number>", methods=["GET", "POST"])
def step(step_number):
    if step_number not in STEPS:
        abort(404)

    step_data = STEPS[step_number]
    if request.method == "POST":
        cv_data = session.get("cv", {})
        for field in step_data["fields"]:
            cv_data[field["name"]] = request.form.get(field["name"], "").strip()
        session["cv"] = cv_data
        next_url = url_for("summary") if step_number == 6 else url_for("step", step_number=step_number + 1)
        return redirect(next_url)

    cv_data = session.get("cv", {})
    fields = [
        {**field, "current_value": cv_data.get(field["name"], field["value"])}
        for field in step_data["fields"]
    ]
    return render_template(
        "form.html",
        step_number=step_number,
        title=step_data["title"],
        next_label=step_data["next_label"],
        fields=fields,
    )


@app.get("/summary")
def summary():
    return render_template("summary.html", fields=SUMMARY_FIELDS, cv=session.get("cv", {}))


@app.get("/download")
def download():
    cv = session.get("cv", {})
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    normal_style = document.styles["Normal"]
    normal_style.font.name = "Aptos"
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(55, 65, 81)

    name = cv.get("name", "").strip() or "Professional CV"
    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(4)
    name_run = title.add_run(name.upper())
    name_run.bold = True
    name_run.font.name = "Aptos Display"
    name_run.font.size = Pt(22)
    name_run.font.color.rgb = RGBColor(25, 76, 80)

    contact_line = " | ".join(
        value for value in (cv.get("email", "").strip(), cv.get("contactInfo", "").strip()) if value
    )
    if contact_line:
        contact = document.add_paragraph()
        contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
        contact.paragraph_format.space_after = Pt(12)
        contact_run = contact.add_run(contact_line)
        contact_run.font.size = Pt(9)
        contact_run.font.color.rgb = RGBColor(91, 101, 111)

    def add_section(heading, entries):
        populated_entries = [(label, value.strip()) for label, value in entries if value and value.strip()]
        if not populated_entries:
            return
        heading_paragraph = document.add_paragraph()
        heading_paragraph.paragraph_format.space_before = Pt(10)
        heading_paragraph.paragraph_format.space_after = Pt(4)
        heading_run = heading_paragraph.add_run(heading.upper())
        heading_run.bold = True
        heading_run.font.name = "Aptos Display"
        heading_run.font.size = Pt(11)
        heading_run.font.color.rgb = RGBColor(25, 76, 80)
        for label, value in populated_entries:
            paragraph = document.add_paragraph()
            paragraph.paragraph_format.space_after = Pt(3)
            if label:
                label_run = paragraph.add_run(f"{label}: ")
                label_run.bold = True
            paragraph.add_run(value)

    add_section("Professional Summary", [("", cv.get("bio", ""))])
    add_section("Education", [
        ("Degree", cv.get("degree", "")),
        ("Institution", cv.get("institution", "")),
        ("Graduation year", cv.get("year", "")),
    ])
    add_section("Work Experience", [
        ("Job title", cv.get("jobTitle", "")),
        ("Company", cv.get("company", "")),
        ("Duration", cv.get("duration", "")),
    ])
    add_section("Skills", [("", cv.get("skills", ""))])
    add_section("Projects", [("", cv.get("projects", ""))])

    output = BytesIO()
    document.save(output)
    output.seek(0)
    filename = "_".join(name.split()) or "professional_cv"
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=f"{filename}_CV.docx",
    )


@app.get("/assets/<path:filename>")
def assets(filename):
    if filename not in {"1.jpg", "background2.jpg"}:
        abort(404)
    return send_from_directory(app.root_path, filename)


if __name__ == "__main__":
    app.run(debug=True)