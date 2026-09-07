from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    send_file,
    send_from_directory,
)

from werkzeug.security import (
    generate_password_hash,
    check_password_hash,
)

from werkzeug.utils import secure_filename

from database import (
    get_db_connection,
    init_db,
    upgrade_database,
)

from openpyxl import Workbook

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import (
    getSampleStyleSheet,
    ParagraphStyle,
)
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

import os
import random
import html
import requests
import base64

from datetime import (
    datetime,
    timedelta,
    timezone,
)

from email.message import EmailMessage
from email.utils import make_msgid


# ==================================================
# FLASK APPLICATION
# ==================================================

app = Flask(__name__)

APPLICATION_STATUSES = [
    "Submitted",
    "Under Review",
    "Shortlisted",
    "Selected",
    "Rejected",
]


# ==================================================
# SESSION CONFIGURATION
# ==================================================

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "change-this-secret-key-before-production",
)

app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = (
    os.environ.get("SESSION_COOKIE_SECURE", "false").lower() == "true"
)
app.permanent_session_lifetime = timedelta(days=1)


# ==================================================
# BREVO SMTP CONFIGURATION
# ==================================================

SMTP_LOGIN = os.environ.get("SMTP_LOGIN", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp-relay.brevo.com")

try:
    SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
except (TypeError, ValueError):
    SMTP_PORT = 587

FROM_EMAIL = os.environ.get(
    "FROM_EMAIL",
    "trekso275@gmail.com",
)
FROM_NAME = os.environ.get(
    "FROM_NAME",
    "Trekso Careers",
)

BREVO_API_KEY = os.environ.get("BREVO_API_KEY", "").strip()


# ==================================================
# ADMIN CONFIGURATION
# ==================================================

ADMIN_USERNAME = os.environ.get(
    "ADMIN_USERNAME",
    "admin",
)
ADMIN_PASSWORD = os.environ.get(
    "ADMIN_PASSWORD",
    "trekso1245",
)


# ==================================================
# WALK-IN INTERVIEW CONFIGURATION
# ==================================================

WALKIN_OFFICE_ADDRESS = os.environ.get(
    "WALKIN_OFFICE_ADDRESS",
    "Trekso Office Address - To Be Updated",
)

WALKIN_DATE = os.environ.get(
    "WALKIN_DATE",
    "Interview Date - To Be Updated",
)

WALKIN_TIME = os.environ.get(
    "WALKIN_TIME",
    "Interview Time - To Be Updated",
)

WALKIN_CONTACT = os.environ.get(
    "WALKIN_CONTACT",
    "Contact Number - To Be Updated",
)


# ==================================================
# FILE CONFIGURATION
# ==================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
PDF_FOLDER = os.path.join(BASE_DIR, "generated_pdfs")

ALLOWED_EXTENSIONS = {
    "pdf",
    "doc",
    "docx",
}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PDF_FOLDER, exist_ok=True)


# ==================================================
# DATABASE INITIALIZATION
# ==================================================

try:
    init_db()
    upgrade_database()
    print("Database initialization completed.")
except Exception:
    app.logger.exception("Database initialization failed during startup.")


# ==================================================
# HELPERS
# ==================================================

def allowed_file(filename):
    return (
        bool(filename)
        and "." in filename
        and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS
    )


def utc_now_naive():
    """Return current UTC time without timezone information."""
    return datetime.now(timezone.utc).replace(tzinfo=None)


def get_current_user():
    user_id = session.get("user_id")

    if not user_id:
        return None

    connection = get_db_connection()

    try:
        return connection.execute(
            """
            SELECT *
            FROM users
            WHERE id = %s
            """,
            (user_id,),
        ).fetchone()
    finally:
        connection.close()


def application_exists_for_user(user_id):
    connection = get_db_connection()

    try:
        return connection.execute(
            """
            SELECT id
            FROM applications
            WHERE user_id = %s
            LIMIT 1
            """,
            (user_id,),
        ).fetchone()
    finally:
        connection.close()


def get_application_by_id(application_id):
    connection = get_db_connection()

    try:
        return connection.execute(
            """
            SELECT *
            FROM applications
            WHERE id = %s
            """,
            (application_id,),
        ).fetchone()
    finally:
        connection.close()


def applicant_logged_in():
    return bool(session.get("user_id"))


def admin_logged_in():
    return bool(session.get("admin_logged_in"))


def application_pdf_path(application_id):
    return os.path.join(
        PDF_FOLDER,
        f"application_{application_id}.pdf",
    )


def applicant_login_required():
    if not session.get("user_id"):
        return False

    user = get_current_user()

    if not user:
        session.clear()
        return False

    if int(user.get("email_verified", 0) or 0) != 1:
        session.clear()
        return False

    return True


# ==================================================
# SEND EMAIL
# ==================================================
# ==================================================
# SEND EMAIL THROUGH BREVO HTTPS API
# ==================================================

def send_email(
    recipient_email,
    subject,
    body,
    attachment_path=None,
    attachment_name=None,
):
    """
    Send Trekso email through Brevo HTTPS API only.

    Returns:
        True  -> email sent successfully
        False -> email sending failed
    """

    brevo_api_key = os.environ.get(
        "BREVO_API_KEY",
        ""
    ).strip()

    if not brevo_api_key:
        app.logger.error(
            "BREVO_API_KEY is not configured."
        )
        return False

    if not FROM_EMAIL:
        app.logger.error(
            "FROM_EMAIL is not configured."
        )
        return False

    if not recipient_email:
        app.logger.error(
            "Recipient email is empty."
        )
        return False

    # --------------------------------------------------
    # Trekso logo
    # --------------------------------------------------

    logo_url = os.environ.get(
        "LOGO_URL",
        "https://trekso-hiring-website.onrender.com/static/images/trekso-logo.png",
    )

    safe_body = html.escape(
        str(body)
    ).replace(
        "\n",
        "<br>"
    )

    safe_subject = html.escape(
        str(subject)
    )

    safe_logo_url = html.escape(
        logo_url,
        quote=True,
    )

    # --------------------------------------------------
    # HTML EMAIL
    # --------------------------------------------------

    html_body = f"""
<!DOCTYPE html>
<html lang="en">

<head>

    <meta charset="UTF-8">

    <meta
        name="viewport"
        content="width=device-width, initial-scale=1.0"
    >

    <title>{safe_subject}</title>

</head>

<body style="
    margin:0;
    padding:0;
    background:#f5f5f5;
    font-family:Arial,Helvetica,sans-serif;
">

    <div style="
        max-width:620px;
        margin:30px auto;
        background:#ffffff;
        border-radius:14px;
        overflow:hidden;
        box-shadow:0 5px 25px rgba(0,0,0,0.08);
    ">

        <div style="
            padding:25px 20px;
            text-align:center;
            border-bottom:1px solid #eeeeee;
            background:#ffffff;
        ">

            <img
                src="{safe_logo_url}"
                alt="Trekso"
                style="
                    max-width:190px;
                    width:100%;
                    height:auto;
                    display:block;
                    margin:0 auto;
                "
            >

        </div>

        <div style="
            padding:30px;
            color:#333333;
            font-size:15px;
            line-height:1.7;
        ">

            {safe_body}

        </div>

        <div style="
            padding:18px;
            text-align:center;
            background:#111111;
            color:#ffffff;
            font-size:12px;
        ">

            © 2026 Trekso Careers.
            All rights reserved.

        </div>

    </div>

</body>
</html>
"""

    # --------------------------------------------------
    # BREVO API PAYLOAD
    # --------------------------------------------------

    payload = {
        "sender": {
            "name": FROM_NAME,
            "email": FROM_EMAIL,
        },

        "to": [
            {
                "email": recipient_email,
            }
        ],

        "subject": subject,

        "textContent": body,

        "htmlContent": html_body,
    }

    # --------------------------------------------------
    # PDF ATTACHMENT
    # --------------------------------------------------

    if attachment_path:

        if not os.path.isfile(
            attachment_path
        ):
            app.logger.error(
                "Email attachment does not exist: %s",
                attachment_path,
            )
            return False

        try:

            with open(
                attachment_path,
                "rb"
            ) as attachment_file:

                encoded_file = base64.b64encode(
                    attachment_file.read()
                ).decode(
                    "utf-8"
                )

            payload["attachment"] = [
                {
                    "content": encoded_file,
                    "name": (
                        attachment_name
                        or os.path.basename(
                            attachment_path
                        )
                    ),
                }
            ]

        except Exception:

            app.logger.exception(
                "Could not prepare email attachment."
            )

            return False

    # --------------------------------------------------
    # BREVO HTTPS API
    # --------------------------------------------------

    try:

        response = requests.post(
            "https://api.brevo.com/v3/smtp/email",

            headers={
                "accept": "application/json",
                "api-key": brevo_api_key,
                "content-type": "application/json",
            },

            json=payload,

            timeout=15,
        )

        # --------------------------------------------------
        # SUCCESS
        # --------------------------------------------------

        if 200 <= response.status_code < 300:

            app.logger.info(
                "Brevo API email sent successfully to %s",
                recipient_email,
            )

            return True

        # --------------------------------------------------
        # BREVO ERROR
        # --------------------------------------------------

        app.logger.error(
            "Brevo API rejected email. "
            "Status=%s Response=%s",
            response.status_code,
            response.text,
        )

        return False

    # --------------------------------------------------
    # TIMEOUT
    # --------------------------------------------------

    except requests.Timeout:

        app.logger.error(
            "Brevo API request timed out."
        )

        return False

    # --------------------------------------------------
    # NETWORK ERROR
    # --------------------------------------------------

    except requests.RequestException:

        app.logger.exception(
            "Network error while connecting to Brevo API."
        )

        return False

    # --------------------------------------------------
    # UNEXPECTED ERROR
    # --------------------------------------------------

    except Exception:

        app.logger.exception(
            "Unexpected error while sending email."
        )

        return False
# ==================================================
# SEND OTP EMAIL
# ==================================================

def send_otp_email(recipient_email, otp):
    subject = "Trekso - Email Verification OTP"

    body = f"""
Dear Applicant,

Welcome to Trekso Careers.

Your email verification OTP is:

{otp}

This OTP is valid for 10 minutes.

Please do not share this OTP with anyone.

Regards,
Trekso Careers
"""

    return send_email(
        recipient_email,
        subject,
        body,
    )


# ==================================================
# CREATE APPLICATION PDF
# ==================================================

def create_application_pdf(application):
    pdf_filename = f"application_{application['id']}.pdf"
    pdf_path = os.path.join(PDF_FOLDER, pdf_filename)

    document = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Title"],
        fontSize=20,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#111111"),
        spaceAfter=10,
    )

    subtitle_style = ParagraphStyle(
        "SubtitleStyle",
        parent=styles["Heading2"],
        alignment=TA_CENTER,
        textColor=colors.HexColor("#ff4b00"),
        spaceAfter=20,
    )

    normal_style = ParagraphStyle(
        "NormalStyle",
        parent=styles["Normal"],
        fontSize=10,
        leading=15,
    )

    footer_style = ParagraphStyle(
        "FooterStyle",
        parent=normal_style,
        alignment=TA_CENTER,
        fontSize=8,
        textColor=colors.grey,
    )

    def pdf_text(value, default=""):
        if value is None:
            value = default
        return html.escape(str(value)).replace("\n", "<br/>")

    story = []

    story.append(Paragraph("TREKSO", title_style))
    story.append(Paragraph("JOB APPLICATION", subtitle_style))

    story.append(
        Paragraph(
            f"<b>Application ID:</b> #{pdf_text(application['id'])}",
            normal_style,
        )
    )

    story.append(Spacer(1, 8))

    story.append(
        Paragraph(
            f"<b>Application Status:</b> {pdf_text(application.get('status', 'Submitted'))}",
            normal_style,
        )
    )

    story.append(Spacer(1, 20))

    data = [
        ["Field", "Details"],
        ["Job Position", pdf_text(application.get("position"))],
        ["Full Name", pdf_text(application.get("name"))],
        ["Email", pdf_text(application.get("email"))],
        ["Mobile Number", pdf_text(application.get("mobile"))],
        ["Date of Birth", pdf_text(application.get("dob"))],
        ["Gender", pdf_text(application.get("gender"))],
        ["Qualification", pdf_text(application.get("qualification"))],
        ["College Name", pdf_text(application.get("college_name"))],
        ["University Name", pdf_text(application.get("university_name"))],
        ["Branch / Specialization", pdf_text(application.get("branch"))],
        ["Graduation Year", pdf_text(application.get("graduation_year"))],
        ["Candidate Type", pdf_text(application.get("candidate_type"))],
        ["Experience", pdf_text(application.get("experience"), "N/A")],
        ["Technical Skills", pdf_text(application.get("skills"))],
        ["Address", pdf_text(application.get("address"))],
        ["Resume", pdf_text(application.get("resume_filename"), "Not uploaded")],
        ["Submitted On", pdf_text(application.get("created_at"))],
    ]

    table = Table(
        data,
        colWidths=[160, 330],
        repeatRows=1,
    )

    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#ff4b00")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )

    story.append(table)
    story.append(Spacer(1, 25))

    story.append(
        Paragraph(
            "Please carry this application copy when attending the Trekso walk-in interview.",
            normal_style,
        )
    )

    story.append(Spacer(1, 15))

    story.append(
        Paragraph(
            f"<b>Walk-in Interview Address:</b> {pdf_text(WALKIN_OFFICE_ADDRESS)}",
            normal_style,
        )
    )
    story.append(
        Paragraph(
            f"<b>Interview Date:</b> {pdf_text(WALKIN_DATE)}",
            normal_style,
        )
    )
    story.append(
        Paragraph(
            f"<b>Interview Time:</b> {pdf_text(WALKIN_TIME)}",
            normal_style,
        )
    )
    story.append(
        Paragraph(
            f"<b>Contact:</b> {pdf_text(WALKIN_CONTACT)}",
            normal_style,
        )
    )

    story.append(Spacer(1, 20))

    story.append(
        Paragraph(
            "This document was generated electronically by Trekso Careers.",
            footer_style,
        )
    )

    document.build(story)

    return pdf_path


# ==================================================
# APPLICATION SUCCESS EMAIL
# ==================================================

def send_application_success_email(application, pdf_path):
    subject = "Trekso - Application Submitted Successfully"

    body = f"""
Dear {application['name']},

Congratulations!

Your job application has been successfully submitted to Trekso Careers.

APPLICATION DETAILS
-------------------

Application ID:
#{application['id']}

Position:
{application['position']}

Status:
{application['status']}

Name:
{application['name']}

Email:
{application['email']}

Mobile:
{application['mobile']}

Your submitted application PDF is attached to this email.

WALK-IN INTERVIEW
-----------------

Please attend the Trekso walk-in interview at our office.

Office Address:
{WALKIN_OFFICE_ADDRESS}

Interview Date:
{WALKIN_DATE}

Interview Time:
{WALKIN_TIME}

Contact Number:
{WALKIN_CONTACT}

Please carry the application PDF and your resume when attending the interview.

Regards,

Trekso Careers
"""

    return send_email(
        application["email"],
        subject,
        body,
        attachment_path=pdf_path,
        attachment_name=(
            f"Trekso_Application_{application['id']}.pdf"
        ),
    )


# ==================================================
# HOME PAGE
# ==================================================

@app.route("/")
def home():
    return render_template("home.html")


# Compatibility endpoint for older error.html files using url_for('index').
@app.route("/index")
def index():
    return redirect(url_for("home"))


# ==================================================
# APPLICANT REGISTRATION
# ==================================================

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "GET":
        return render_template("register.html")

    name = request.form.get("name", "").strip()
    email = request.form.get("email", "").strip().lower()
    mobile = request.form.get("mobile", "").strip()
    password = request.form.get("password", "")

    if not name or not email or not mobile or not password:
        return render_template(
            "register.html",
            error="All fields are required.",
        )

    connection = get_db_connection()
    otp = None

    try:
        existing_user = connection.execute(
            """
            SELECT id
            FROM users
            WHERE LOWER(email) = LOWER(%s)
            """,
            (email,),
        ).fetchone()

        if existing_user:
            return render_template(
                "register.html",
                error="Email already registered. Please login.",
            )

        existing_mobile_user = connection.execute(
            """
            SELECT id
            FROM users
            WHERE mobile = %s
            """,
            (mobile,),
        ).fetchone()

        if existing_mobile_user:
            return render_template(
                "register.html",
                error="Mobile number already registered.",
            )

        # Remove an old pending registration for the same email so that
        # the newest OTP/password is the active one.
        connection.execute(
            """
            DELETE FROM pending_registrations
            WHERE LOWER(email) = LOWER(%s)
            """,
            (email,),
        ).close()

        pending_mobile = connection.execute(
            """
            SELECT email
            FROM pending_registrations
            WHERE mobile = %s
            """,
            (mobile,),
        ).fetchone()

        if pending_mobile:
            return render_template(
                "register.html",
                error="This mobile number is already under verification.",
            )

        otp = str(random.randint(100000, 999999))

        otp_expiry = (
            utc_now_naive()
            + timedelta(minutes=10)
        )

        password_hash = generate_password_hash(password)

        connection.execute(
            """
            INSERT INTO pending_registrations (
                name,
                email,
                mobile,
                password_hash,
                otp,
                otp_expiry,
                otp_attempts
            )
            VALUES (
                %s,
                %s,
                %s,
                %s,
                %s,
                %s,
                %s
            )
            """,
            (
                name,
                email,
                mobile,
                password_hash,
                otp,
                otp_expiry,
                0,
            ),
        ).close()

        connection.commit()

    except Exception:
        connection.rollback()
        app.logger.exception("Registration database error.")
        return render_template(
            "register.html",
            error="Unable to start registration. Please try again.",
        )

    finally:
        connection.close()

    # Send OTP after the pending registration has been committed.
    email_sent = send_otp_email(email, otp)

    if not email_sent:
        # Remove the pending record so a new registration attempt can
        # generate a fresh OTP instead of leaving a dead pending record.
        cleanup_connection = get_db_connection()

        try:
            cleanup_connection.execute(
                """
                DELETE FROM pending_registrations
                WHERE LOWER(email) = LOWER(%s)
                """,
                (email,),
            ).close()
            cleanup_connection.commit()
        except Exception:
            cleanup_connection.rollback()
            app.logger.exception("Could not clean failed pending registration.")
        finally:
            cleanup_connection.close()

        return render_template(
            "register.html",
            error=(
                "Unable to send verification email right now. "
                "Please try again in a few minutes."
            ),
        )

    session.clear()
    session.permanent = True
    session["pending_registration_email"] = email

    return redirect(url_for("verify_email"))


# ==================================================
# VERIFY EMAIL OTP
# ==================================================

@app.route("/verify-email", methods=["GET", "POST"])
def verify_email():
    email = session.get("pending_registration_email")

    if not email:
        return redirect(url_for("register"))

    connection = get_db_connection()

    try:
        pending = connection.execute(
            """
            SELECT *
            FROM pending_registrations
            WHERE LOWER(email) = LOWER(%s)
            """,
            (email,),
        ).fetchone()

        if not pending:
            session.pop("pending_registration_email", None)
            return render_template(
                "register.html",
                error=(
                    "Registration session expired. Please register again."
                ),
            )

        if request.method == "GET":
            return render_template("verify_email.html")

        entered_otp = request.form.get("otp", "").strip()

        if not entered_otp:
            return render_template(
                "verify_email.html",
                error="Please enter the OTP.",
            )

        attempts = int(pending["otp_attempts"] or 0)

        if attempts >= 5:
            return render_template(
                "verify_email.html",
                error=(
                    "Too many incorrect OTP attempts. Please register again."
                ),
            )

        expiry = pending["otp_expiry"]

        if expiry.tzinfo is None:
            expiry = expiry.replace(tzinfo=timezone.utc)

        if datetime.now(timezone.utc) > expiry:
            return render_template(
                "verify_email.html",
                error="OTP expired. Please register again.",
            )

        if entered_otp != str(pending["otp"]):
            attempts += 1

            connection.execute(
                """
                UPDATE pending_registrations
                SET otp_attempts = %s
                WHERE id = %s
                """,
                (attempts, pending["id"]),
            ).close()

            connection.commit()

            remaining = max(0, 5 - attempts)

            return render_template(
                "verify_email.html",
                error=(
                    f"Incorrect OTP. {remaining} attempt(s) remaining."
                ),
            )

        existing_user = connection.execute(
            """
            SELECT id
            FROM users
            WHERE LOWER(email) = LOWER(%s)
            """,
            (pending["email"],),
        ).fetchone()

        if existing_user:
            session.pop("pending_registration_email", None)
            return render_template(
                "login.html",
                error=(
                    "This email is already registered. Please login."
                ),
            )

        existing_mobile = connection.execute(
            """
            SELECT id
            FROM users
            WHERE mobile = %s
            """,
            (pending["mobile"],),
        ).fetchone()

        if existing_mobile:
            session.pop("pending_registration_email", None)
            return render_template(
                "register.html",
                error=(
                    "This mobile number is already registered."
                ),
            )

        connection.execute(
            """
            INSERT INTO users (
                name,
                email,
                mobile,
                password,
                email_verified,
                mobile_verified
            )
            VALUES (
                %s,
                %s,
                %s,
                %s,
                %s,
                %s
            )
            """,
            (
                pending["name"],
                pending["email"],
                pending["mobile"],
                pending["password_hash"],
                1,
                0,
            ),
        ).close()

        connection.execute(
            """
            DELETE FROM pending_registrations
            WHERE id = %s
            """,
            (pending["id"],),
        ).close()

        connection.commit()

        session.pop("pending_registration_email", None)

        return render_template(
            "login.html",
            success=(
                "Email verified successfully. Your Trekso account has been created. "
                "You can now login."
            ),
        )

    except Exception:
        connection.rollback()
        app.logger.exception("OTP verification error.")
        return render_template(
            "verify_email.html",
            error="Unable to verify OTP right now. Please try again.",
        ), 500

    finally:
        connection.close()


# ==================================================
# APPLICANT LOGIN
# ==================================================

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET":
        return render_template("login.html")

    email = request.form.get("email", "").strip().lower()
    password = request.form.get("password", "")

    connection = get_db_connection()

    try:
        user = connection.execute(
            """
            SELECT *
            FROM users
            WHERE LOWER(email) = LOWER(%s)
            """,
            (email,),
        ).fetchone()
    finally:
        connection.close()

    valid_password = False

    if user and user.get("password"):
        try:
            valid_password = check_password_hash(
                user["password"],
                password,
            )
        except (TypeError, ValueError):
            valid_password = False

    if valid_password:
        if int(user.get("email_verified", 0) or 0) != 1:
            return render_template(
                "login.html",
                error="Please verify your email before logging in.",
            )

        # Clear any admin/pending session and establish a fresh applicant
        # session. This makes switching between different applicants in
        # the same browser deterministic.
        session.clear()
        session.permanent = True
        session["user_id"] = user["id"]
        session["user_name"] = user["name"]
        session["user_email"] = user["email"]
        session["user_mobile"] = user["mobile"]

        return redirect(url_for("applicant_home"))

    return render_template(
        "login.html",
        error="Invalid email or password.",
    )


# ==================================================
# APPLICANT DASHBOARD
# ==================================================

@app.route("/applicant-home")
def applicant_home():
    if not applicant_login_required():
        return redirect(url_for("login"))

    user = get_current_user()
    application = application_exists_for_user(user["id"])

    if application:
        application = get_application_by_id(application["id"])

    return render_template(
        "applicant_home.html",
        application=application,
    )


# ==================================================
# JOB APPLICATION
# ==================================================

@app.route("/apply", methods=["GET", "POST"])
def apply():
    if not applicant_login_required():
        return redirect(url_for("login"))

    user = get_current_user()

    existing_application = application_exists_for_user(user["id"])

    if existing_application:
        safe_position = html.escape(
            str(
                get_application_by_id(existing_application["id"])["position"]
            )
        )

        return f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Application Already Submitted</title>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    background: #fff7f2;
                    text-align: center;
                    padding-top: 100px;
                }}
                .box {{
                    background: white;
                    width: 500px;
                    max-width: 90%;
                    margin: auto;
                    padding: 40px;
                    border-radius: 14px;
                    box-shadow: 0 10px 35px rgba(0,0,0,0.08);
                }}
                h2 {{ color: #ff4b00; margin-bottom: 20px; }}
                p {{ color: #555; margin-bottom: 15px; }}
                a {{
                    display: inline-block;
                    background: #ff4b00;
                    color: white;
                    text-decoration: none;
                    padding: 12px 20px;
                    border-radius: 7px;
                }}
            </style>
        </head>
        <body>
            <div class="box">
                <h2>Application Already Submitted</h2>
                <p>You have already applied for:</p>
                <p><strong>{safe_position}</strong></p>
                <p>One applicant can submit only one application.</p>
                <a href="/my-application">View My Application</a>
            </div>
        </body>
        </html>
        """

    if request.method == "GET":
        return render_template("apply.html")

    position = request.form.get("position", "").strip()
    name = request.form.get("name", "").strip()
    email = request.form.get("email", "").strip().lower()
    mobile = request.form.get("mobile", "").strip()
    dob = request.form.get("dob", "").strip()
    gender = request.form.get("gender", "").strip()
    qualification = request.form.get("qualification", "").strip()
    college_name = request.form.get("college_name", "").strip()
    university_name = request.form.get("university_name", "").strip()
    branch = request.form.get("branch", "").strip()
    graduation_year = request.form.get("graduation_year", "").strip()
    candidate_type = request.form.get("candidate_type", "Fresher").strip()
    experience = request.form.get("experience", "").strip()
    skills = request.form.get("skills", "").strip()
    address = request.form.get("address", "").strip()

    required_fields = {
        "position": position,
        "name": name,
        "email": email,
        "mobile": mobile,
        "dob": dob,
        "gender": gender,
        "qualification": qualification,
        "college_name": college_name,
        "university_name": university_name,
        "branch": branch,
        "graduation_year": graduation_year,
        "candidate_type": candidate_type,
        "skills": skills,
        "address": address,
    }

    for field_name, field_value in required_fields.items():
        if not field_value:
            return render_template(
                "apply.html",
                error=(
                    field_name.replace("_", " ").title()
                    + " is required."
                ),
            )

    resume = request.files.get("resume")
    saved_resume_filename = None
    resume_path = None

    connection = get_db_connection()

    try:
        current_user = connection.execute(
            """
            SELECT *
            FROM users
            WHERE id = %s
            """,
            (session["user_id"],),
        ).fetchone()

        if not current_user:
            session.clear()
            return redirect(url_for("login"))

        if int(current_user.get("email_verified", 0) or 0) != 1:
            return render_template(
                "login.html",
                error="Please verify your email before applying.",
            )

        existing_application = connection.execute(
            """
            SELECT id
            FROM applications
            WHERE user_id = %s
            LIMIT 1
            """,
            (session["user_id"],),
        ).fetchone()

        if existing_application:
            return redirect(url_for("my_application"))

        existing_email_application = connection.execute(
            """
            SELECT id
            FROM applications
            WHERE LOWER(email) = LOWER(%s)
            LIMIT 1
            """,
            (email,),
        ).fetchone()

        if existing_email_application:
            return render_template(
                "apply.html",
                error=(
                    "This email has already been used to submit an application."
                ),
            )

        existing_mobile_application = connection.execute(
            """
            SELECT id
            FROM applications
            WHERE mobile = %s
            LIMIT 1
            """,
            (mobile,),
        ).fetchone()

        if existing_mobile_application:
            return render_template(
                "apply.html",
                error=(
                    "This mobile number has already been used to submit an application."
                ),
            )

        if resume and resume.filename:
            if not allowed_file(resume.filename):
                return render_template(
                    "apply.html",
                    error=(
                        "Invalid resume format. Only PDF, DOC and DOCX files are allowed."
                    ),
                )

            original_filename = secure_filename(resume.filename)

            if not original_filename:
                return render_template(
                    "apply.html",
                    error="Invalid resume filename.",
                )

            base_name, extension = os.path.splitext(original_filename)

            saved_resume_filename = (
                f"{session['user_id']}_{base_name}{extension}"
            )

            resume_path = os.path.join(
                UPLOAD_FOLDER,
                saved_resume_filename,
            )

            resume.save(resume_path)

        inserted_row = connection.execute(
            """
            INSERT INTO applications (
                user_id,
                position,
                name,
                email,
                mobile,
                dob,
                gender,
                qualification,
                college_name,
                university_name,
                branch,
                graduation_year,
                candidate_type,
                experience,
                skills,
                address,
                resume_filename,
                status
            )
            VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s, %s, %s, 'Submitted'
            )
            RETURNING id
            """,
            (
                session["user_id"],
                position,
                name,
                email,
                mobile,
                dob,
                gender,
                qualification,
                college_name,
                university_name,
                branch,
                graduation_year,
                candidate_type,
                experience,
                skills,
                address,
                saved_resume_filename,
            ),
        ).fetchone()

        application_id = inserted_row["id"]
        connection.commit()

    except Exception:
        connection.rollback()

        if resume_path and os.path.isfile(resume_path):
            try:
                os.remove(resume_path)
            except OSError:
                pass

        app.logger.exception("Application submission database error.")
        return render_template(
            "apply.html",
            error="Unable to submit your application right now. Please try again.",
        ), 500

    finally:
        connection.close()

    application = get_application_by_id(application_id)

    try:
        pdf_path = create_application_pdf(application)
    except Exception:
        app.logger.exception("Application PDF generation failed.")
        return render_template(
            "error.html",
            error="Application was saved, but the application PDF could not be generated.",
        ), 500

    email_sent = send_application_success_email(
        application,
        pdf_path,
    )

    email_message = (
        "<p>A confirmation email containing your application PDF has been sent to your registered email address.</p>"
        if email_sent
        else
        "<p style=\"color:#b00020;\">Your application was saved successfully, but the confirmation email could not be sent. Please download the application PDF from your dashboard.</p>"
    )

    safe_name = html.escape(str(application["name"]))
    safe_address = html.escape(str(WALKIN_OFFICE_ADDRESS))
    safe_date = html.escape(str(WALKIN_DATE))
    safe_time = html.escape(str(WALKIN_TIME))
    safe_contact = html.escape(str(WALKIN_CONTACT))

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Trekso | Application Submitted</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                background: #fff7f2;
                text-align: center;
                padding-top: 70px;
            }}
            .box {{
                background: white;
                width: 580px;
                max-width: 92%;
                margin: auto;
                padding: 45px;
                border-radius: 14px;
                box-shadow: 0 10px 35px rgba(0,0,0,0.08);
            }}
            h2 {{ color: #ff4b00; margin-bottom: 20px; }}
            p {{ color: #555; margin-bottom: 16px; line-height: 1.6; }}
            .application-id {{
                font-size: 20px;
                font-weight: bold;
                color: #111;
                margin: 20px 0;
            }}
            .walkin {{
                background: #fff7f2;
                padding: 20px;
                border-radius: 10px;
                margin: 25px 0;
                text-align: left;
            }}
            .walkin h3 {{ color: #ff4b00; margin-top: 0; }}
            a {{
                display: inline-block;
                background: #ff4b00;
                color: white;
                text-decoration: none;
                padding: 12px 22px;
                border-radius: 7px;
                font-weight: bold;
            }}
        </style>
    </head>
    <body>
        <div class="box">
            <h2>Application Submitted Successfully!</h2>
            <p>Dear {safe_name},</p>
            <p>Your Trekso job application has been successfully submitted.</p>
            <div class="application-id">Application ID: #{application['id']}</div>
            {email_message}
            <div class="walkin">
                <h3>Walk-in Interview</h3>
                <p><strong>Office Address:</strong><br>{safe_address}</p>
                <p><strong>Date:</strong><br>{safe_date}</p>
                <p><strong>Time:</strong><br>{safe_time}</p>
                <p><strong>Contact:</strong><br>{safe_contact}</p>
                <p>Please carry your application PDF and resume when attending the interview.</p>
            </div>
            <a href="/applicant-home">Go to Dashboard</a>
        </div>
    </body>
    </html>
    """


# ==================================================
# VIEW MY APPLICATION
# ==================================================

@app.route("/my-application")
def my_application():
    if not applicant_login_required():
        return redirect(url_for("login"))

    user = get_current_user()
    application_ref = application_exists_for_user(user["id"])

    if not application_ref:
        return render_template(
            "my_application.html",
            application=None,
        )

    application = get_application_by_id(application_ref["id"])

    return render_template(
        "my_application.html",
        application=application,
    )


# ==================================================
# APPLICANT DOWNLOAD RESUME
# ==================================================

@app.route("/download/resume/<int:application_id>")
def download_resume(application_id):
    if not applicant_login_required():
        return redirect(url_for("login"))

    application = get_application_by_id(application_id)

    if not application:
        return "Application not found.", 404

    if application["user_id"] != session["user_id"]:
        return "Access denied.", 403

    resume_filename = application.get("resume_filename")

    if not resume_filename:
        return "No resume was uploaded.", 404

    resume_path = os.path.join(
        UPLOAD_FOLDER,
        resume_filename,
    )

    if not os.path.isfile(resume_path):
        return "Resume file not found.", 404

    return send_from_directory(
        UPLOAD_FOLDER,
        resume_filename,
        as_attachment=True,
    )


@app.route("/download-resume/<int:application_id>")
def download_resume_legacy(application_id):
    return download_resume(application_id)


# ==================================================
# DOWNLOAD APPLICATION PDF
# ==================================================

@app.route("/download/application/<int:application_id>")
def download_application_by_id(application_id):
    """
    Download a specific application PDF for the logged-in applicant.
    The application must belong to the current applicant.
    """
    if not applicant_login_required():
        return redirect(url_for("login"))

    application = get_application_by_id(application_id)

    if not application:
        return "Application not found.", 404

    if application["user_id"] != session["user_id"]:
        return "Access denied.", 403

    pdf_path = application_pdf_path(application_id)

    if not os.path.isfile(pdf_path):
        pdf_path = create_application_pdf(application)

    return send_file(
        pdf_path,
        as_attachment=True,
        download_name=(
            f"Trekso_Application_{application_id}.pdf"
        ),
    )


# Main endpoint used by the applicant dashboard/templates.
# It finds the current applicant's application automatically.
@app.route("/download-application")
def download_application():
    if not applicant_login_required():
        return redirect(url_for("login"))

    application_ref = application_exists_for_user(
        session["user_id"]
    )

    if not application_ref:
        return "No application found.", 404

    return download_application_by_id(
        application_ref["id"]
    )


# Legacy ID-based route kept for compatibility.
@app.route("/download-application/<int:application_id>")
def download_application_legacy(application_id):
    return download_application_by_id(
        application_id
    )


# Compatibility endpoint for templates/code that use
# url_for("download_my_application").
@app.route("/download-my-application")
def download_my_application():
    return download_application()


# ==================================================
# ADMIN LOGIN
# ==================================================

@app.route("/admin", methods=["GET", "POST"])
def admin():
    return redirect(url_for("admin_login"))


@app.route("/admin-login", methods=["GET", "POST"])
def admin_login():
    if request.method == "GET":
        return render_template("admin_login.html")

    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")

    if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
        session.clear()
        session.permanent = True
        session["admin_logged_in"] = True
        return redirect(url_for("admin_dashboard"))

    return render_template(
        "admin_login.html",
        error="Invalid admin username or password.",
    )


# ==================================================
# ADMIN DASHBOARD
# ==================================================

@app.route("/admin-dashboard")
def admin_dashboard():
    if not admin_logged_in():
        return redirect(url_for("admin_login"))

    search = request.args.get("search", "").strip()
    status_filter = request.args.get("status", "").strip()

    connection = get_db_connection()

    try:
        total_count = connection.execute(
            """
            SELECT COUNT(*) AS count
            FROM applications
            """
        ).fetchone()["count"]

        submitted_count = connection.execute(
            """
            SELECT COUNT(*) AS count
            FROM applications
            WHERE status = 'Submitted'
            """
        ).fetchone()["count"]

        under_review_count = connection.execute(
            """
            SELECT COUNT(*) AS count
            FROM applications
            WHERE status = 'Under Review'
            """
        ).fetchone()["count"]

        shortlisted_count = connection.execute(
            """
            SELECT COUNT(*) AS count
            FROM applications
            WHERE status = 'Shortlisted'
            """
        ).fetchone()["count"]

        selected_count = connection.execute(
            """
            SELECT COUNT(*) AS count
            FROM applications
            WHERE status = 'Selected'
            """
        ).fetchone()["count"]

        rejected_count = connection.execute(
            """
            SELECT COUNT(*) AS count
            FROM applications
            WHERE status = 'Rejected'
            """
        ).fetchone()["count"]

        query = """
            SELECT
                id,
                user_id,
                position,
                name,
                email,
                mobile,
                dob,
                gender,
                qualification,
                college_name,
                university_name,
                branch,
                graduation_year,
                candidate_type,
                experience,
                skills,
                address,
                resume_filename,
                status,
                created_at
            FROM applications
            WHERE 1 = 1
        """

        params = []

        if search:
            query += """
                AND (
                    name ILIKE %s
                    OR email ILIKE %s
                    OR mobile ILIKE %s
                    OR position ILIKE %s
                    OR qualification ILIKE %s
                    OR college_name ILIKE %s
                    OR university_name ILIKE %s
                    OR branch ILIKE %s
                    OR skills ILIKE %s
                )
            """

            search_value = f"%{search}%"
            params.extend([search_value] * 9)

        if status_filter in APPLICATION_STATUSES:
            query += " AND status = %s"
            params.append(status_filter)

        query += " ORDER BY created_at DESC, id DESC"

        applications = connection.execute(
            query,
            tuple(params),
        ).fetchall()

        return render_template(
            "admin_dashboard.html",
            applications=applications,
            total_count=total_count,
            submitted_count=submitted_count,
            under_review_count=under_review_count,
            shortlisted_count=shortlisted_count,
            selected_count=selected_count,
            rejected_count=rejected_count,
            application_statuses=APPLICATION_STATUSES,
            search=search,
            status_filter=status_filter,
        )

    except Exception:
        app.logger.exception("Admin dashboard error.")
        return render_template(
            "error.html",
            error="Unable to load admin dashboard.",
        ), 500

    finally:
        connection.close()


# ==================================================
# ADMIN VIEW APPLICATION
# ==================================================

@app.route("/admin-application/<int:application_id>")
def admin_view_application(application_id):
    if not admin_logged_in():
        return redirect(url_for("admin_login"))

    application = get_application_by_id(application_id)

    if not application:
        return "Application not found.", 404

    return render_template(
        "admin_application.html",
        application=application,
    )


# Compatibility alias.
@app.route("/admin/application/<int:application_id>")
def admin_view_application_legacy(application_id):
    return admin_view_application(application_id)


# ==================================================
# ADMIN UPDATE STATUS
# ==================================================

@app.route("/admin-update-status/<int:application_id>", methods=["POST"])
def update_application_status(application_id):
    if not admin_logged_in():
        return redirect(url_for("admin_login"))

    new_status = request.form.get("status", "").strip()

    if new_status not in APPLICATION_STATUSES:
        return "Invalid application status.", 400

    connection = get_db_connection()

    try:
        application = connection.execute(
            """
            SELECT id
            FROM applications
            WHERE id = %s
            """,
            (application_id,),
        ).fetchone()

        if not application:
            return "Application not found.", 404

        connection.execute(
            """
            UPDATE applications
            SET status = %s
            WHERE id = %s
            """,
            (new_status, application_id),
        ).close()

        connection.commit()

    except Exception:
        connection.rollback()
        app.logger.exception("Application status update failed.")
        return "Unable to update application status.", 500

    finally:
        connection.close()

    return redirect(
        url_for(
            "admin_view_application",
            application_id=application_id,
        )
    )


@app.route("/admin/application/<int:application_id>/status", methods=["POST"])
def update_application_status_legacy(application_id):
    return update_application_status(application_id)


# ==================================================
# ADMIN DELETE APPLICATION
# ==================================================

@app.route("/admin-delete-application/<int:application_id>", methods=["POST"])
def admin_delete_application(application_id):
    if not admin_logged_in():
        return redirect(url_for("admin_login"))

    connection = get_db_connection()
    resume_filename = None

    try:
        application = connection.execute(
            """
            SELECT id, resume_filename
            FROM applications
            WHERE id = %s
            """,
            (application_id,),
        ).fetchone()

        if not application:
            return "Application not found.", 404

        resume_filename = application.get("resume_filename")

        connection.execute(
            """
            DELETE FROM applications
            WHERE id = %s
            """,
            (application_id,),
        ).close()

        connection.commit()

    except Exception:
        connection.rollback()
        app.logger.exception("Application deletion failed.")
        return "Unable to delete application.", 500

    finally:
        connection.close()

    if resume_filename:
        resume_path = os.path.join(
            UPLOAD_FOLDER,
            resume_filename,
        )

        if os.path.isfile(resume_path):
            try:
                os.remove(resume_path)
            except OSError:
                pass

    pdf_path = application_pdf_path(application_id)

    if os.path.isfile(pdf_path):
        try:
            os.remove(pdf_path)
        except OSError:
            pass

    return redirect(url_for("admin_dashboard"))


@app.route("/admin/application/<int:application_id>/delete", methods=["POST"])
def admin_delete_application_legacy(application_id):
    return admin_delete_application(application_id)


# ==================================================
# ADMIN DOWNLOAD RESUME
# ==================================================

@app.route("/admin-download-resume/<int:application_id>")
def admin_download_resume(application_id):
    if not admin_logged_in():
        return redirect(url_for("admin_login"))

    application = get_application_by_id(application_id)

    if not application:
        return "Application not found.", 404

    resume_filename = application.get("resume_filename")

    if not resume_filename:
        return "No resume was uploaded.", 404

    resume_path = os.path.join(
        UPLOAD_FOLDER,
        resume_filename,
    )

    if not os.path.isfile(resume_path):
        return "Resume file not found.", 404

    return send_from_directory(
        UPLOAD_FOLDER,
        resume_filename,
        as_attachment=True,
    )


@app.route("/admin/download/resume/<int:application_id>")
def admin_download_resume_legacy(application_id):
    return admin_download_resume(application_id)


# ==================================================
# ADMIN DOWNLOAD APPLICATION PDF
# ==================================================

@app.route("/admin/download/application/<int:application_id>")
def admin_download_application(application_id):
    if not admin_logged_in():
        return redirect(url_for("admin_login"))

    application = get_application_by_id(application_id)

    if not application:
        return "Application not found.", 404

    pdf_path = application_pdf_path(application_id)

    if not os.path.isfile(pdf_path):
        pdf_path = create_application_pdf(application)

    return send_file(
        pdf_path,
        as_attachment=True,
        download_name=f"Trekso_Application_{application_id}.pdf",
    )


@app.route("/admin/application/<int:application_id>/download")
def admin_download_application_legacy(application_id):
    return admin_download_application(application_id)


# ==================================================
# ADMIN EXPORT TO EXCEL
# ==================================================

@app.route("/admin-export-excel")
def admin_export_excel():
    if not admin_logged_in():
        return redirect(url_for("admin_login"))

    connection = get_db_connection()

    try:
        applications = connection.execute(
            """
            SELECT *
            FROM applications
            ORDER BY id DESC
            """
        ).fetchall()
    finally:
        connection.close()

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Applications"

    headings = [
        "Application ID",
        "Position",
        "Name",
        "Email",
        "Mobile",
        "Date of Birth",
        "Gender",
        "Qualification",
        "College Name",
        "University Name",
        "Branch",
        "Graduation Year",
        "Candidate Type",
        "Experience",
        "Skills",
        "Address",
        "Resume",
        "Status",
        "Submitted On",
    ]

    worksheet.append(headings)

    for application in applications:
        worksheet.append(
            [
                application.get("id"),
                application.get("position"),
                application.get("name"),
                application.get("email"),
                application.get("mobile"),
                application.get("dob"),
                application.get("gender"),
                application.get("qualification"),
                application.get("college_name"),
                application.get("university_name"),
                application.get("branch"),
                application.get("graduation_year"),
                application.get("candidate_type"),
                application.get("experience"),
                application.get("skills"),
                application.get("address"),
                application.get("resume_filename"),
                application.get("status"),
                application.get("created_at"),
            ]
        )

    column_widths = {
        "A": 15,
        "B": 20,
        "C": 25,
        "D": 30,
        "E": 18,
        "F": 15,
        "G": 12,
        "H": 20,
        "I": 38,
        "J": 30,
        "K": 25,
        "L": 18,
        "M": 18,
        "N": 25,
        "O": 35,
        "P": 40,
        "Q": 30,
        "R": 18,
        "S": 22,
    }

    for column, width in column_widths.items():
        worksheet.column_dimensions[column].width = width

    worksheet.freeze_panes = "A2"

    excel_path = os.path.join(
        PDF_FOLDER,
        "applications.xlsx",
    )

    workbook.save(excel_path)

    return send_file(
        excel_path,
        as_attachment=True,
        download_name="applications.xlsx",
    )


@app.route("/admin/export")
def admin_export_alias():
    return admin_export_excel()


@app.route("/admin/export-excel")
def admin_export_excel_alias():
    return admin_export_excel()


@app.route("/export-excel")
def export_excel_alias():
    return admin_export_excel()


# ==================================================
# ADMIN LOGOUT
# ==================================================

@app.route("/admin-logout")
def admin_logout():
    session.clear()
    return redirect(url_for("admin_login"))


# ==================================================
# APPLICANT LOGOUT
# ==================================================

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ==================================================
# ERROR HANDLERS
# ==================================================

@app.errorhandler(413)
def file_too_large(error):
    return (
        """
        <h2 style="text-align:center;margin-top:100px;color:#ff4b00;">
            Resume file is too large.
        </h2>
        <p style="text-align:center;">
            Maximum allowed file size is 5 MB.
        </p>
        <p style="text-align:center;">
            <a href="/apply">Go Back</a>
        </p>
        """,
        413,
    )


@app.errorhandler(500)
def internal_server_error(error):
    app.logger.exception("Unhandled internal server error.")

    try:
        return render_template(
            "error.html",
            error="An internal server error occurred. Please try again.",
        ), 500
    except Exception:
        return (
            "An internal server error occurred. Please try again.",
            500,
        )


# ==================================================
# RUN APPLICATION
# ==================================================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False,
    )
