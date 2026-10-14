
from flask import Flask, render_template, request, send_file
from pathlib import Path
from io import BytesIO
from datetime import date
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
)

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "employees.xlsx"

LEAVE_TYPES = [
    "Vacation Leave",
    "Sick Leave",
    "Emergency Leave",
    "Personal Leave",
    "Other"
]


def load_employees():
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            "employees.xlsx is missing from the project folder."
        )

    df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")

    required = ["EmployeeID", "FullName", "Position", "Department"]
    missing = [column for column in required if column not in df.columns]

    if missing:
        raise ValueError(
            "Missing Excel columns: " + ", ".join(missing)
        )

    df["EmployeeID"] = df["EmployeeID"].str.strip()
    return df


def create_leave_pdf(employee, leave_type, start, end, reason):
    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=45,
        rightMargin=45,
        topMargin=40,
        bottomMargin=40
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "FormTitle",
        parent=styles["Title"],
        fontSize=16,
        leading=20,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#17365D")
    )

    story = [
        Paragraph("EMPLOYEE LEAVE APPLICATION", title_style),
        Spacer(1, 8),
        Paragraph(
            "Application date: " + date.today().strftime("%B %d, %Y"),
            styles["Normal"]
        ),
        Spacer(1, 22),
        Paragraph("A. EMPLOYEE INFORMATION", styles["Heading2"]),
        Spacer(1, 8)
    ]

    employee_rows = [
        ["Employee ID", escape(employee["EmployeeID"])],
        ["Full Name", escape(employee["FullName"])],
        ["Position", escape(employee["Position"])],
        ["Department", escape(employee["Department"])]
    ]

    employee_table = Table(employee_rows, colWidths=[130, 340])
    employee_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.grey),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#EAF0F8")),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9)
    ]))

    story.extend([
        employee_table,
        Spacer(1, 22),
        Paragraph("B. LEAVE DETAILS", styles["Heading2"]),
        Spacer(1, 8)
    ])

    days = (end - start).days + 1

    leave_rows = [
        ["Type of Leave", escape(leave_type)],
        ["Start Date", start.strftime("%B %d, %Y")],
        ["End Date", end.strftime("%B %d, %Y")],
        ["Calendar Days", str(days)],
        ["Reason", escape(reason)]
    ]

    leave_table = Table(leave_rows, colWidths=[130, 340])
    leave_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.grey),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#EAF0F8")),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9)
    ]))

    story.extend([
        leave_table,
        Spacer(1, 35),
        Paragraph(
            "Applicant's Signature: ______________________________",
            styles["Normal"]
        ),
        Spacer(1, 30),
        Paragraph(
            "Recommendation:  [ ] Approved  [ ] Disapproved",
            styles["Normal"]
        ),
        Spacer(1, 28),
        Paragraph(
            "Approving Officer: __________________________________",
            styles["Normal"]
        ),
        Spacer(1, 24),
        Paragraph(
            "This is a leave application, not confirmation of approval.",
            styles["Italic"]
        )
    ])

    doc.build(story)
    buffer.seek(0)
    return buffer


@app.route("/", methods=["GET", "POST"])
def index():
    error = None

    try:
        employees = load_employees()
        employee_ids = employees["EmployeeID"].tolist()
    except (OSError, ValueError, ImportError) as exc:
        employees = pd.DataFrame()
        employee_ids = []
        error = str(exc)

    if request.method == "POST" and error is None:
        employee_id = request.form.get("employee_id", "").strip()
        leave_type = request.form.get("leave_type", "").strip()
        start_text = request.form.get("start_date", "").strip()
        end_text = request.form.get("end_date", "").strip()
        reason = request.form.get("reason", "").strip()

        try:
            if leave_type not in LEAVE_TYPES:
                raise ValueError("Please select a valid leave type.")

            if not reason:
                raise ValueError("Please enter the reason for your leave.")

            start = date.fromisoformat(start_text)
            end = date.fromisoformat(end_text)

            if start < date.today():
                raise ValueError("The start date cannot be in the past.")

            if end < start:
                raise ValueError("The end date cannot be before the start date.")

            matches = employees[
                employees["EmployeeID"] == employee_id
            ]

            if not employee_id or matches.empty:
                raise ValueError("Please select a valid employee.")

            employee = matches.iloc[0].to_dict()

            pdf = create_leave_pdf(
                employee, leave_type, start, end, reason
            )

            return send_file(
                pdf,
                mimetype="application/pdf",
                as_attachment=True,
                download_name=f"Leave_Application_{employee_id}.pdf"
            )

        except ValueError as exc:
            error = str(exc)

    return render_template(
        "index.html",
        employee_ids=employee_ids,
        leave_types=LEAVE_TYPES,
        error=error
    )


if __name__ == "__main__":
    app.run(debug=True)
