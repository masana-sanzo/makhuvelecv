import os
import secrets
from io import BytesIO

from flask import Flask, abort, redirect, render_template, request, send_file, send_from_directory, session, url_for

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
    text = "\n".join(f"{label}: {cv.get(key, '')}" for label, key in SUMMARY_FIELDS)
    return send_file(
        BytesIO(text.encode("utf-8")),
        mimetype="text/plain",
        as_attachment=True,
        download_name="summary.txt",
    )


@app.get("/assets/<path:filename>")
def assets(filename):
    if filename not in {"1.jpg", "background2.jpg"}:
        abort(404)
    return send_from_directory(app.root_path, filename)


if __name__ == "__main__":
    app.run(debug=True)