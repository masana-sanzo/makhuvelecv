import os
import secrets
from io import BytesIO
from xml.sax.saxutils import escape

from flask import Flask, abort, redirect, render_template, request, send_file, send_from_directory, session, url_for
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or secrets.token_hex(32)

STEPS = {
    1: {
        "title": "Personal Information",
        "description": "Start with the contact details employers need and a short introduction to your professional strengths.",
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
        "description": "Add each qualification separately so employers can see your degree, institution and graduation date clearly.",
        "next_label": "Next Page: Work Experience",
        "kind": "education",
    },
    3: {
        "title": "Work Experience",
        "description": "Add each role separately. Focus responsibilities on what you did and achievements on the results you delivered.",
        "next_label": "Next Page: Skills",
        "kind": "experience",
    },
    4: {
        "title": "Skills",
        "description": "List relevant skills one at a time. Choose specific abilities that match the roles you are targeting.",
        "next_label": "Next: Projects",
        "kind": "skills",
    },
    5: {
        "title": "Projects",
        "description": "Show selected projects that demonstrate your practical experience, initiative or specialist knowledge.",
        "next_label": "Next: Contact",
        "fields": [
            {"name": "projects", "label": "Projects (comma-separated)", "value": "Project 1, Project 2, Project 3"},
        ],
    },
    6: {
        "title": "Additional Information",
        "description": "Add any relevant extras, then choose a CV style that suits your career stage and application.",
        "next_label": "Review CV",
        "kind": "additional",
        "fields": [
            {"name": "contactInfo", "label": "Contact information", "type": "textarea", "value": ""},
            {"name": "certifications", "label": "Certifications (one per line)", "type": "textarea", "value": ""},
            {"name": "languages", "label": "Languages (one per line)", "type": "textarea", "value": ""},
            {"name": "volunteering", "label": "Volunteering", "type": "textarea", "value": ""},
            {"name": "achievements", "label": "Achievements (one per line)", "type": "textarea", "value": ""},
            {"name": "references", "label": "Professional references", "type": "textarea", "value": ""},
        ],
    },
}

TEMPLATES = {
    "professional": {
        "label": "Professional",
        "description": "A clean, straightforward, ATS-friendly layout for most industries and online applications.",
        "audience": "Best for corporate, administration, finance, technology and experienced professional roles.",
        "color": "194C50",
        "font": "Aptos",
    },
    "modern": {
        "label": "Modern",
        "description": "A contemporary layout with stronger visual hierarchy and a more distinctive presentation.",
        "audience": "Best for creative, communications, product and other roles where a polished visual style helps.",
        "color": "2855A6",
        "font": "Aptos",
    },
    "academic": {
        "label": "Academic",
        "description": "A formal, text-focused layout that gives qualifications and scholarly experience room to lead.",
        "audience": "Best for students, graduates, teachers, researchers and academic or education applications.",
        "color": "5C4635",
        "font": "Georgia",
    },
}

ASSET_FILES = {
    "Thoughtful Professional at Her Desk-1.png",
    "Graduation Success in a Modern Study-2.png",
    "Collaborative Office Teamwork-3.png",
    "Professional Achievement Planning Desk-4.png",
    "Confident Resume Presentation in a Modern Office-5.png",
}


def form_list(field_name):
    return [value.strip() for value in request.form.getlist(field_name) if value.strip()]


def build_cv_content(cv):
    sections = []
    if cv.get("bio", "").strip():
        sections.append(("Professional Summary", [cv["bio"].strip()]))
    if cv.get("education"):
        sections.append(("Education", [
            " | ".join(value for value in (item.get("degree"), item.get("institution"), item.get("year")) if value)
            for item in cv["education"]
        ]))
    if cv.get("experience"):
        entries = []
        for item in cv["experience"]:
            heading = " | ".join(value for value in (item.get("job_title"), item.get("employer"), item.get("dates")) if value)
            detail = "\n".join(value for value in (item.get("responsibilities"), item.get("achievements")) if value)
            entries.append("\n".join(value for value in (heading, detail) if value))
        sections.append(("Work Experience", entries))
    for key, heading in (
        ("skills", "Skills"),
        ("projects", "Projects"),
        ("certifications", "Certifications"),
        ("languages", "Languages"),
        ("volunteering", "Volunteering"),
        ("achievements", "Achievements"),
        ("references", "Professional References"),
    ):
        values = cv.get(key, []) if isinstance(cv.get(key), list) else cv.get(key, "").splitlines()
        values = [value.strip() for value in values if value.strip()]
        if values:
            sections.append((heading, values))
    return sections


@app.get("/")
def index():
    return render_template("index.html")


@app.route("/step/<int:step_number>", methods=["GET", "POST"])
def step(step_number):
    if step_number not in STEPS:
        abort(404)

    step_data = STEPS[step_number]
    if request.method == "POST":
        cv_data = dict(session.get("cv", {}))
        if step_number == 2:
            degrees = request.form.getlist("education_degree[]")
            institutions = request.form.getlist("education_institution[]")
            years = request.form.getlist("education_year[]")
            cv_data["education"] = [
                {"degree": degree.strip(), "institution": institution.strip(), "year": year.strip()}
                for degree, institution, year in zip(degrees, institutions, years)
                if degree.strip() or institution.strip() or year.strip()
            ]
        elif step_number == 3:
            job_titles = request.form.getlist("experience_job_title[]")
            employers = request.form.getlist("experience_employer[]")
            dates = request.form.getlist("experience_dates[]")
            responsibilities = request.form.getlist("experience_responsibilities[]")
            achievements = request.form.getlist("experience_achievements[]")
            cv_data["experience"] = [
                {
                    "job_title": job_title.strip(),
                    "employer": employer.strip(),
                    "dates": date_range.strip(),
                    "responsibilities": responsibility.strip(),
                    "achievements": achievement.strip(),
                }
                for job_title, employer, date_range, responsibility, achievement in zip(
                    job_titles, employers, dates, responsibilities, achievements
                )
                if any(value.strip() for value in (job_title, employer, date_range, responsibility, achievement))
            ]
        elif step_number == 4:
            cv_data["skills"] = form_list("skills[]")
        else:
            for field in step_data.get("fields", []):
                value = request.form.get(field["name"], "").strip()
                cv_data[field["name"]] = value
            if step_number == 6:
                cv_data["template"] = request.form.get("template", "professional")
                if cv_data["template"] not in TEMPLATES:
                    cv_data["template"] = "professional"
        session["cv"] = cv_data
        next_url = url_for("summary") if step_number == 6 else url_for("step", step_number=step_number + 1)
        return redirect(next_url)

    cv_data = session.get("cv", {})
    fields = [
        {**field, "current_value": cv_data.get(field["name"], field["value"])}
        for field in step_data.get("fields", [])
    ]
    return render_template(
        "form.html",
        step_number=step_number,
        title=step_data["title"],
        next_label=step_data["next_label"],
        fields=fields,
        kind=step_data.get("kind", "standard"),
        step_description=step_data["description"],
        cv=cv_data,
        templates=TEMPLATES,
    )


@app.get("/summary")
def summary():
    cv = session.get("cv", {})
    return render_template(
        "summary.html",
        cv=cv,
        sections=build_cv_content(cv),
        templates=TEMPLATES,
    )


def build_word_document(cv, selected_template):
    style = TEMPLATES[selected_template]
    accent = RGBColor.from_string(style["color"])
    academic = selected_template == "academic"
    modern = selected_template == "modern"
    document = Document()
    page = document.sections[0]
    page.top_margin = Inches(0.65)
    page.bottom_margin = Inches(0.65)
    page.left_margin = Inches(0.8)
    page.right_margin = Inches(0.8)
    normal = document.styles["Normal"]
    normal.font.name = style["font"]
    normal.font.size = Pt(10 if not academic else 11)
    normal.font.color.rgb = RGBColor(55, 65, 81)

    name = cv.get("name", "").strip() or "Professional CV"
    heading = document.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER if academic or modern else WD_ALIGN_PARAGRAPH.LEFT
    name_run = heading.add_run(name.upper() if not academic else name)
    name_run.bold = True
    name_run.font.name = style["font"]
    name_run.font.size = Pt(24 if modern else 20)
    name_run.font.color.rgb = accent
    contact_line = " | ".join(value for value in (cv.get("email", "").strip(), cv.get("contactInfo", "").strip()) if value)
    if contact_line:
        contact = document.add_paragraph(contact_line)
        contact.alignment = WD_ALIGN_PARAGRAPH.CENTER if academic or modern else WD_ALIGN_PARAGRAPH.LEFT

    for section_name, entries in build_cv_content(cv):
        section_heading = document.add_paragraph()
        section_heading.paragraph_format.space_before = Pt(10)
        section_heading.paragraph_format.space_after = Pt(4)
        section_run = section_heading.add_run(section_name if academic else section_name.upper())
        section_run.bold = True
        section_run.font.name = style["font"]
        section_run.font.size = Pt(12 if academic else 11)
        section_run.font.color.rgb = accent
        for entry in entries:
            paragraph = document.add_paragraph()
            paragraph.paragraph_format.space_after = Pt(4)
            if modern or academic:
                paragraph.style = document.styles["List Bullet"]
            paragraph.add_run(entry)
    return document


def build_pdf(cv, selected_template):
    palette = {
        "professional": colors.HexColor("#194C50"),
        "modern": colors.HexColor("#2855A6"),
        "academic": colors.HexColor("#5C4635"),
    }
    accent = palette[selected_template]
    academic = selected_template == "academic"
    modern = selected_template == "modern"
    name = cv.get("name", "").strip() or "Professional CV"
    output = BytesIO()
    pdf = SimpleDocTemplate(
        output,
        pagesize=letter,
        rightMargin=0.75 * inch,
        leftMargin=0.75 * inch,
        topMargin=0.65 * inch,
        bottomMargin=0.65 * inch,
        title=f"{name} CV",
        author=name,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CVTitle", parent=styles["Title"], fontName="Times-Bold" if academic else "Helvetica-Bold",
        fontSize=22 if modern else 20, leading=26, textColor=accent,
        alignment=TA_CENTER if academic or modern else TA_LEFT, spaceAfter=5,
    )
    contact_style = ParagraphStyle(
        "CVContact", parent=styles["Normal"], fontSize=9, textColor=colors.HexColor("#5B656F"),
        alignment=TA_CENTER if academic or modern else TA_LEFT, spaceAfter=12,
    )
    section_style = ParagraphStyle(
        "CVSection", parent=styles["Heading2"], fontName="Times-Bold" if academic else "Helvetica-Bold",
        fontSize=12 if academic else 11, leading=15, textColor=accent, spaceBefore=10, spaceAfter=4,
    )
    body_style = ParagraphStyle(
        "CVBody", parent=styles["BodyText"], fontName="Times-Roman" if academic else "Helvetica",
        fontSize=10 if not academic else 11, leading=14, textColor=colors.HexColor("#374151"), spaceAfter=4,
        leftIndent=12 if modern or academic else 0,
    )
    story = [Paragraph(escape(name.upper() if not academic else name), title_style)]
    contacts = " | ".join(value for value in (cv.get("email", "").strip(), cv.get("contactInfo", "").strip()) if value)
    if contacts:
        story.append(Paragraph(escape(contacts), contact_style))
    for section_name, entries in build_cv_content(cv):
        story.append(Paragraph(escape(section_name if academic else section_name.upper()), section_style))
        story.extend(Paragraph(escape(entry).replace("\n", "<br/>"), body_style) for entry in entries)
    pdf.build(story)
    output.seek(0)
    return output


def selected_cv_template(cv):
    selected = cv.get("template", "professional")
    return selected if selected in TEMPLATES else "professional"


@app.get("/download/<format_name>")
def download(format_name):
    if format_name not in {"word", "pdf"}:
        abort(404)
    cv = session.get("cv", {})
    selected = selected_cv_template(cv)
    name = cv.get("name", "").strip() or "professional_cv"
    filename = "_".join(name.split()) or "professional_cv"
    if format_name == "word":
        output = BytesIO()
        build_word_document(cv, selected).save(output)
        output.seek(0)
        mimetype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        extension = "docx"
    else:
        output = build_pdf(cv, selected)
        mimetype = "application/pdf"
        extension = "pdf"
    return send_file(
        output,
        mimetype=mimetype,
        as_attachment=True,
        download_name=f"{filename}_CV.{extension}",
    )


@app.get("/assets/<path:filename>")
def assets(filename):
    if filename not in ASSET_FILES:
        abort(404)
    return send_from_directory(os.path.join(app.root_path, "assets"), filename)


if __name__ == "__main__":
    app.run(debug=True)