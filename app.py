from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    send_file,
    send_from_directory
)

from werkzeug.security import (
    generate_password_hash,
    check_password_hash
)

from werkzeug.utils import secure_filename

from database import (
    get_db_connection,
    init_db,
    upgrade_database
)

from openpyxl import Workbook

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import (
    getSampleStyleSheet,
    ParagraphStyle
)
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)

import os
import smtplib
import random
import html

from datetime import (
    datetime,
    timedelta,
    timezone
)

from email.message import EmailMessage
from email.utils import make_msgid


# ==================================================
# FLASK APPLICATION
# ==================================================

app = Flask(__name__)


# ==================================================
# SESSION CONFIGURATION
# ==================================================

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "change-this-secret-key-before-production"
)

app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = False

app.permanent_session_lifetime = 60 * 60 * 24


# ==================================================
# BREVO SMTP CONFIGURATION
# ==================================================

SMTP_LOGIN = os.environ.get(
    "SMTP_LOGIN",
    ""
)

SMTP_PASSWORD = os.environ.get(
    "SMTP_PASSWORD",
    ""
)

SMTP_HOST = os.environ.get(
    "SMTP_HOST",
    "smtp-relay.brevo.com"
)

SMTP_PORT = int(
    os.environ.get(
        "SMTP_PORT",
        "587"
    )
)

FROM_EMAIL = os.environ.get(
    "FROM_EMAIL",
    "trekso275@gmail.com"
)

FROM_NAME = os.environ.get(
    "FROM_NAME",
    "Trekso Careers"
)


# ==================================================
# WALK-IN INTERVIEW CONFIGURATION
# ==================================================

WALKIN_OFFICE_ADDRESS = os.environ.get(
    "WALKIN_OFFICE_ADDRESS",
    "Trekso Office Address - To Be Updated"
)

WALKIN_DATE = os.environ.get(
    "WALKIN_DATE",
    "Interview Date - To Be Updated"
)

WALKIN_TIME = os.environ.get(
    "WALKIN_TIME",
    "Interview Time - To Be Updated"
)

WALKIN_CONTACT = os.environ.get(
    "WALKIN_CONTACT",
    "Contact Number - To Be Updated"
)


# ==================================================
# FILE CONFIGURATION
# ==================================================

UPLOAD_FOLDER = "uploads"

PDF_FOLDER = "generated_pdfs"

ALLOWED_EXTENSIONS = {
    "pdf",
    "doc",
    "docx"
}


app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

app.config["MAX_CONTENT_LENGTH"] = (
    5 * 1024 * 1024
)


os.makedirs(
    UPLOAD_FOLDER,
    exist_ok=True
)

os.makedirs(
    PDF_FOLDER,
    exist_ok=True
)


# ==================================================
# DATABASE
# ==================================================

init_db()
upgrade_database()


# ==================================================
# HELPER
# ==================================================

def allowed_file(filename):

    return (
        "." in filename
        and filename.rsplit(
            ".",
            1
        )[1].lower()
        in ALLOWED_EXTENSIONS
    )


# ==================================================
# SEND EMAIL
# ==================================================

def send_email(
    recipient_email,
    subject,
    body,
    attachment_path=None,
    attachment_name=None
):

    if not SMTP_LOGIN:

        raise RuntimeError(
            "SMTP_LOGIN is not configured."
        )

    if not SMTP_PASSWORD:

        raise RuntimeError(
            "SMTP_PASSWORD is not configured."
        )

    if not FROM_EMAIL:

        raise RuntimeError(
            "FROM_EMAIL is not configured."
        )


    message = EmailMessage()


    message["Subject"] = subject


    message["From"] = (
        f"{FROM_NAME} <{FROM_EMAIL}>"
    )


    message["To"] = recipient_email


    # ------------------------------------------------
    # Plain text
    # ------------------------------------------------

    message.set_content(
        body
    )


    # ------------------------------------------------
    # Trekso logo
    # ------------------------------------------------

    logo_path = os.path.join(
        "static",
        "images",
        "trekso-logo.png"
    )


    logo_cid = make_msgid(
        domain="trekso"
    )


    safe_body = html.escape(
        body
    )


    safe_body = safe_body.replace(
        "\n",
        "<br>"
    )


    html_body = f"""
<!DOCTYPE html>

<html lang="en">

<head>

    <meta charset="UTF-8">

    <meta name="viewport"
          content="width=device-width, initial-scale=1.0">

    <title>{html.escape(subject)}</title>

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
                src="cid:{logo_cid[1:-1]}"
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


    message.add_alternative(
        html_body,
        subtype="html"
    )


    # ------------------------------------------------
    # Inline logo
    # ------------------------------------------------

    if os.path.isfile(
        logo_path
    ):

        with open(
            logo_path,
            "rb"
        ) as logo_file:

            logo_data = logo_file.read()


        html_part = message.get_payload()[-1]


        html_part.add_related(
            logo_data,
            maintype="image",
            subtype="png",
            cid=logo_cid,
            filename="trekso-logo.png"
        )


    # ------------------------------------------------
    # PDF attachment
    # ------------------------------------------------

    if attachment_path:

        with open(
            attachment_path,
            "rb"
        ) as attachment_file:

            file_data = attachment_file.read()


        message.add_attachment(
            file_data,
            maintype="application",
            subtype="pdf",
            filename=(
                attachment_name
                or os.path.basename(
                    attachment_path
                )
            )
        )


    # ------------------------------------------------
    # Send through Brevo
    # ------------------------------------------------

    with smtplib.SMTP(
        SMTP_HOST,
        SMTP_PORT
    ) as server:

        server.starttls()

        server.login(
            SMTP_LOGIN,
            SMTP_PASSWORD
        )

        server.send_message(
            message
        )


# ==================================================
# SEND OTP
# ==================================================

def send_otp_email(
    recipient_email,
    otp
):

    subject = (
        "Trekso - Email Verification OTP"
    )


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


    send_email(
        recipient_email,
        subject,
        body
    )


# ==================================================
# CREATE APPLICATION PDF
# ==================================================

def create_application_pdf(
    application
):

    pdf_filename = (
        "application_"
        + str(application["id"])
        + ".pdf"
    )


    pdf_path = os.path.join(
        PDF_FOLDER,
        pdf_filename
    )


    document = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )


    styles = getSampleStyleSheet()


    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Title"],
        fontSize=20,
        alignment=TA_CENTER,
        textColor=colors.HexColor(
            "#111111"
        ),
        spaceAfter=10
    )


    subtitle_style = ParagraphStyle(
        "SubtitleStyle",
        parent=styles["Heading2"],
        alignment=TA_CENTER,
        textColor=colors.HexColor(
            "#ff4b00"
        ),
        spaceAfter=20
    )


    normal_style = ParagraphStyle(
        "NormalStyle",
        parent=styles["Normal"],
        fontSize=10,
        leading=15
    )


    footer_style = ParagraphStyle(
        "FooterStyle",
        parent=normal_style,
        alignment=TA_CENTER,
        fontSize=8,
        textColor=colors.grey
    )


    story = []


    story.append(
        Paragraph(
            "TREKSO",
            title_style
        )
    )


    story.append(
        Paragraph(
            "JOB APPLICATION",
            subtitle_style
        )
    )


    story.append(
        Paragraph(
            "<b>Application ID:</b> #"
            + str(application["id"]),
            normal_style
        )
    )


    story.append(
        Spacer(
            1,
            8
        )
    )


    story.append(
        Paragraph(
            "<b>Application Status:</b> "
            + str(application["status"]),
            normal_style
        )
    )


    story.append(
        Spacer(
            1,
            20
        )
    )


    data = [

        [
            "Field",
            "Details"
        ],

        [
            "Job Position",
            str(application["position"])
        ],

        [
            "Full Name",
            str(application["name"])
        ],

        [
            "Email",
            str(application["email"])
        ],

        [
            "Mobile Number",
            str(application["mobile"])
        ],

        [
            "Date of Birth",
            str(application["dob"])
        ],

        [
            "Gender",
            str(application["gender"])
        ],

        [
            "Qualification",
            str(application["qualification"])
        ],

        [
            "College Name",
            str(application["college_name"])
        ],

        [
            "University Name",
            str(application["university_name"])
        ],

        [
            "Branch / Specialization",
            str(application["branch"])
        ],

        [
            "Graduation Year",
            str(application["graduation_year"])
        ],

        [
            "Candidate Type",
            str(application["candidate_type"])
        ],

        [
            "Experience",
            str(
                application["experience"]
                or "N/A"
            )
        ],

        [
            "Technical Skills",
            str(application["skills"])
        ],

        [
            "Address",
            str(application["address"])
        ],

        [
            "Resume",
            str(
                application["resume_filename"]
                or "Not uploaded"
            )
        ],

        [
            "Submitted On",
            str(application["created_at"])
        ]

    ]


    table = Table(
        data,
        colWidths=[
            160,
            330
        ],
        repeatRows=1
    )


    table.setStyle(
        TableStyle([

            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor(
                    "#ff4b00"
                )
            ),

            (
                "TEXTCOLOR",
                (0, 0),
                (-1, 0),
                colors.white
            ),

            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold"
            ),

            (
                "FONTNAME",
                (0, 1),
                (0, -1),
                "Helvetica-Bold"
            ),

            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.grey
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP"
            ),

            (
                "FONTSIZE",
                (0, 0),
                (-1, -1),
                9
            ),

            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                8
            ),

            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                8
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                8
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                8
            )

        ])
    )


    story.append(
        table
    )


    story.append(
        Spacer(
            1,
            25
        )
    )


    story.append(
        Paragraph(
            (
                "Please carry this application copy "
                "when attending the Trekso walk-in interview."
            ),
            normal_style
        )
    )


    story.append(
        Spacer(
            1,
            15
        )
    )


    story.append(
        Paragraph(
            (
                "<b>Walk-in Interview Address:</b> "
                + html.escape(
                    str(WALKIN_OFFICE_ADDRESS)
                )
            ),
            normal_style
        )
    )


    story.append(
        Paragraph(
            (
                "<b>Interview Date:</b> "
                + html.escape(
                    str(WALKIN_DATE)
                )
            ),
            normal_style
        )
    )


    story.append(
        Paragraph(
            (
                "<b>Interview Time:</b> "
                + html.escape(
                    str(WALKIN_TIME)
                )
            ),
            normal_style
        )
    )


    story.append(
        Paragraph(
            (
                "<b>Contact:</b> "
                + html.escape(
                    str(WALKIN_CONTACT)
                )
            ),
            normal_style
        )
    )


    story.append(
        Spacer(
            1,
            20
        )
    )


    story.append(
        Paragraph(
            "This document was generated electronically by Trekso Careers.",
            footer_style
        )
    )


    document.build(
        story
    )


    return pdf_path


# ==================================================
# APPLICATION SUCCESS EMAIL
# ==================================================

def send_application_success_email(
    application,
    pdf_path
):

    subject = (
        "Trekso - Application Submitted Successfully"
    )


    body = f"""
Dear {application["name"]},

Congratulations!

Your job application has been successfully
submitted to Trekso Careers.


APPLICATION DETAILS
-------------------

Application ID:
#{application["id"]}

Position:
{application["position"]}

Status:
{application["status"]}

Name:
{application["name"]}

Email:
{application["email"]}

Mobile:
{application["mobile"]}


Your submitted application PDF is attached
to this email.


WALK-IN INTERVIEW
-----------------

Please attend the Trekso walk-in interview
at our office.

Office Address:
{WALKIN_OFFICE_ADDRESS}

Interview Date:
{WALKIN_DATE}

Interview Time:
{WALKIN_TIME}

Contact Number:
{WALKIN_CONTACT}


Please carry the application PDF and your
resume when attending the interview.


Regards,

Trekso Careers
"""


    send_email(
        application["email"],
        subject,
        body,
        attachment_path=pdf_path,
        attachment_name=(
            "Trekso_Application_"
            + str(application["id"])
            + ".pdf"
        )
    )


# ==================================================
# HOME
# ==================================================

@app.route("/")
def home():

    return render_template(
        "home.html"
    )


# ==================================================
# REGISTRATION
#
# IMPORTANT:
# NO user is created here.
# Only pending_registrations is created.
# ==================================================

@app.route(
    "/register",
    methods=["GET", "POST"]
)
def register():

    if request.method == "POST":

        name = request.form.get(
            "name",
            ""
        ).strip()

        email = request.form.get(
            "email",
            ""
        ).strip().lower()

        mobile = request.form.get(
            "mobile",
            ""
        ).strip()

        password = request.form.get(
            "password",
            ""
        )


        # ------------------------------------------
        # Validation
        # ------------------------------------------

        if not name or not email or not mobile or not password:

            return render_template(
                "register.html",
                error="All fields are required."
            )


        connection = get_db_connection()


        # ------------------------------------------
        # Check already registered email
        # ------------------------------------------

        existing_user = connection.execute(
            """
            SELECT id
            FROM users
            WHERE LOWER(email) = LOWER(?)
            """,
            (
                email,
            )
        ).fetchone()


        if existing_user:

            connection.close()

            return render_template(
                "register.html",
                error=(
                    "Email already registered. "
                    "Please login."
                )
            )


        # ------------------------------------------
        # Check already registered mobile
        # ------------------------------------------

        existing_mobile_user = connection.execute(
            """
            SELECT id
            FROM users
            WHERE mobile = ?
            """,
            (
                mobile,
            )
        ).fetchone()


        if existing_mobile_user:

            connection.close()

            return render_template(
                "register.html",
                error=(
                    "Mobile number already registered."
                )
            )


        # ------------------------------------------
        # Existing pending email
        #
        # Delete old unfinished verification so
        # applicant can request a fresh OTP.
        # ------------------------------------------

        pending_email = connection.execute(
            """
            SELECT id
            FROM pending_registrations
            WHERE LOWER(email) = LOWER(?)
            """,
            (
                email,
            )
        ).fetchone()


        if pending_email:

            connection.execute(
                """
                DELETE FROM pending_registrations
                WHERE LOWER(email) = LOWER(?)
                """,
                (
                    email,
                )
            )


        # ------------------------------------------
        # Pending mobile with another email
        # ------------------------------------------

        pending_mobile = connection.execute(
            """
            SELECT email
            FROM pending_registrations
            WHERE mobile = ?
            """,
            (
                mobile,
            )
        ).fetchone()


        if pending_mobile:

            connection.close()

            return render_template(
                "register.html",
                error=(
                    "This mobile number is already "
                    "under verification."
                )
            )


        # ------------------------------------------
        # Generate OTP
        # ------------------------------------------

        otp = str(
            random.randint(
                100000,
                999999
            )
        )


        otp_expiry = (
            datetime.now(
                timezone.utc
            )
            + timedelta(
                minutes=10
            )
        ).isoformat()


        # ------------------------------------------
        # Hash password
        #
        # IMPORTANT:
        # Plain password is never stored.
        # ------------------------------------------

        password_hash = generate_password_hash(
            password
        )


        # ------------------------------------------
        # Save pending registration ONLY
        # ------------------------------------------

        try:

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
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    name,
                    email,
                    mobile,
                    password_hash,
                    otp,
                    otp_expiry,
                    0
                )
            )


            connection.commit()


        except Exception:

            connection.rollback()
            connection.close()

            raise


        connection.close()


        # ------------------------------------------
        # Send OTP
        # ------------------------------------------

        try:

            send_otp_email(
                email,
                otp
            )


        except Exception as error:

            print(
                "OTP email error:",
                error
            )


            connection = get_db_connection()


            connection.execute(
                """
                DELETE FROM pending_registrations
                WHERE LOWER(email) = LOWER(?)
                """,
                (
                    email,
                )
            )


            connection.commit()
            connection.close()


            return render_template(
                "register.html",
                error=(
                    "Unable to send verification email. "
                    "Please check the email configuration "
                    "and try again."
                )
            )


        # ------------------------------------------
        # Store pending email only in session
        # ------------------------------------------

        session.clear()

        session["pending_registration_email"] = (
            email
        )


        return redirect(
            url_for(
                "verify_email"
            )
        )


    return render_template(
        "register.html"
    )


# ==================================================
# VERIFY EMAIL OTP
#
# ONLY HERE IS THE ACTUAL USER ACCOUNT CREATED.
# ==================================================

@app.route(
    "/verify-email",
    methods=["GET", "POST"]
)
def verify_email():

    email = session.get(
        "pending_registration_email"
    )


    if not email:

        return redirect(
            url_for(
                "register"
            )
        )


    connection = get_db_connection()


    pending = connection.execute(
        """
        SELECT *
        FROM pending_registrations
        WHERE LOWER(email) = LOWER(?)
        """,
        (
            email,
        )
    ).fetchone()


    if not pending:

        connection.close()

        session.pop(
            "pending_registration_email",
            None
        )

        return render_template(
            "register.html",
            error=(
                "Registration session expired. "
                "Please register again."
            )
        )


    if request.method == "POST":

        entered_otp = request.form.get(
            "otp",
            ""
        ).strip()


        if not entered_otp:

            connection.close()

            return render_template(
                "verify_email.html",
                error="Please enter the OTP."
            )


        # ------------------------------------------
        # Maximum attempts
        # ------------------------------------------

        attempts = (
            pending["otp_attempts"]
            or 0
        )


        if attempts >= 5:

            connection.close()

            return render_template(
                "verify_email.html",
                error=(
                    "Too many incorrect OTP attempts. "
                    "Please register again."
                )
            )


        # ------------------------------------------
        # OTP expiry
        # ------------------------------------------

        try:

            expiry = datetime.fromisoformat(
                pending["otp_expiry"]
            )

        except (
            TypeError,
            ValueError
        ):

            connection.close()

            return render_template(
                "verify_email.html",
                error=(
                    "Invalid OTP session. "
                    "Please register again."
                )
            )


        if expiry.tzinfo is None:

            expiry = expiry.replace(
                tzinfo=timezone.utc
            )


        if datetime.now(
            timezone.utc
        ) > expiry:

            connection.close()

            return render_template(
                "verify_email.html",
                error=(
                    "OTP expired. "
                    "Please register again."
                )
            )


        # ------------------------------------------
        # OTP check
        # ------------------------------------------

        if entered_otp != str(
            pending["otp"]
        ):

            attempts += 1


            connection.execute(
                """
                UPDATE pending_registrations
                SET otp_attempts = ?
                WHERE id = ?
                """,
                (
                    attempts,
                    pending["id"]
                )
            )


            connection.commit()
            connection.close()


            remaining = 5 - attempts


            return render_template(
                "verify_email.html",
                error=(
                    "Incorrect OTP. "
                    + str(remaining)
                    + " attempt(s) remaining."
                )
            )


        # ------------------------------------------
        # FINAL DUPLICATE CHECK
        # ------------------------------------------

        existing_user = connection.execute(
            """
            SELECT id
            FROM users
            WHERE LOWER(email) = LOWER(?)
            """,
            (
                pending["email"],
            )
        ).fetchone()


        if existing_user:

            connection.close()

            session.pop(
                "pending_registration_email",
                None
            )

            return render_template(
                "login.html",
                error=(
                    "This email is already registered. "
                    "Please login."
                )
            )


        existing_mobile = connection.execute(
            """
            SELECT id
            FROM users
            WHERE mobile = ?
            """,
            (
                pending["mobile"],
            )
        ).fetchone()


        if existing_mobile:

            connection.close()

            return render_template(
                "register.html",
                error=(
                    "This mobile number is already "
                    "registered."
                )
            )


        # ------------------------------------------
        # CREATE ACTUAL USER
        # ------------------------------------------

        try:

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
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    pending["name"],
                    pending["email"],
                    pending["mobile"],
                    pending["password_hash"],
                    1,
                    0
                )
            )


            connection.execute(
                """
                DELETE FROM pending_registrations
                WHERE id = ?
                """,
                (
                    pending["id"],
                )
            )


            connection.commit()


        except Exception:

            connection.rollback()
            connection.close()

            raise


        connection.close()


        session.pop(
            "pending_registration_email",
            None
        )


        return render_template(
            "login.html",
            success=(
                "Email verified successfully. "
                "Your Trekso account has been created. "
                "You can now login."
            )
        )


    connection.close()


    return render_template(
        "verify_email.html"
    )


# ==================================================
# APPLICANT LOGIN
# ==================================================

@app.route(
    "/login",
    methods=["GET", "POST"]
)
def login():

    if request.method == "POST":

        email = request.form.get(
            "email",
            ""
        ).strip().lower()

        password = request.form.get(
            "password",
            ""
        )


        connection = get_db_connection()


        user = connection.execute(
            """
            SELECT *
            FROM users
            WHERE LOWER(email) = LOWER(?)
            """,
            (
                email,
            )
        ).fetchone()


        connection.close()


        if user and check_password_hash(
            user["password"],
            password
        ):

            if user["email_verified"] != 1:

                return render_template(
                    "login.html",
                    error=(
                        "Please verify your email "
                        "before logging in."
                    )
                )


            session.clear()

            session.permanent = True

            session["user_id"] = (
                user["id"]
            )

            session["user_name"] = (
                user["name"]
            )

            session["user_email"] = (
                user["email"]
            )

            session["user_mobile"] = (
                user["mobile"]
            )


            return redirect(
                url_for(
                    "applicant_home"
                )
            )


        return render_template(
            "login.html",
            error="Invalid email or password."
        )


    return render_template(
        "login.html"
    )


# ==================================================
# APPLICANT DASHBOARD
# ==================================================

@app.route(
    "/applicant-home"
)
def applicant_home():

    if "user_id" not in session:

        return redirect(
            url_for(
                "login"
            )
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT *
        FROM applications
        WHERE user_id = ?
        LIMIT 1
        """,
        (
            session["user_id"],
        )
    ).fetchone()


    connection.close()


    return render_template(
        "applicant_home.html",
        application=application
    )


# ==================================================
# JOB APPLICATION
# ONE APPLICANT = ONE APPLICATION
# ==================================================

@app.route(
    "/apply",
    methods=["GET", "POST"]
)
def apply():

    if "user_id" not in session:

        return redirect(
            url_for(
                "login"
            )
        )


    connection = get_db_connection()


    user = connection.execute(
        """
        SELECT *
        FROM users
        WHERE id = ?
        """,
        (
            session["user_id"],
        )
    ).fetchone()


    if not user:

        connection.close()

        session.clear()

        return redirect(
            url_for(
                "login"
            )
        )


    if user["email_verified"] != 1:

        connection.close()

        return render_template(
            "login.html",
            error=(
                "Please verify your email "
                "before applying."
            )
        )


    existing_application = connection.execute(
        """
        SELECT id, position
        FROM applications
        WHERE user_id = ?
        LIMIT 1
        """,
        (
            session["user_id"],
        )
    ).fetchone()


    connection.close()


    if existing_application:

        return f"""
        <!DOCTYPE html>

        <html>

        <head>

            <title>
                Application Already Submitted
            </title>

            <style>

                body {{
                    font-family:Arial,sans-serif;
                    background:#fff7f2;
                    text-align:center;
                    padding-top:100px;
                }}

                .box {{
                    background:white;
                    width:500px;
                    max-width:90%;
                    margin:auto;
                    padding:40px;
                    border-radius:14px;
                    box-shadow:
                        0 10px 35px
                        rgba(0,0,0,0.08);
                }}

                h2 {{
                    color:#ff4b00;
                }}

                a {{
                    display:inline-block;
                    background:#ff4b00;
                    color:white;
                    text-decoration:none;
                    padding:12px 20px;
                    border-radius:7px;
                }}

            </style>

        </head>

        <body>

            <div class="box">

                <h2>
                    Application Already Submitted
                </h2>

                <p>
                    You have already applied for:
                </p>

                <p>
                    <strong>
                        {html.escape(
                            str(
                                existing_application["position"]
                            )
                        )}
                    </strong>
                </p>

                <p>
                    One applicant can submit only one application.
                </p>

                <a href="/my-application">
                    View My Application
                </a>

            </div>

        </body>

        </html>
        """


    # ==================================================
    # PROCESS APPLICATION
    # ==================================================

    if request.method == "POST":

        position = request.form.get(
            "position",
            ""
        ).strip()

        name = request.form.get(
            "name",
            ""
        ).strip()

        email = request.form.get(
            "email",
            ""
        ).strip().lower()

        mobile = request.form.get(
            "mobile",
            ""
        ).strip()

        dob = request.form.get(
            "dob",
            ""
        ).strip()

        gender = request.form.get(
            "gender",
            ""
        ).strip()

        qualification = request.form.get(
            "qualification",
            ""
        ).strip()

        college_name = request.form.get(
            "college_name",
            ""
        ).strip()

        university_name = request.form.get(
            "university_name",
            ""
        ).strip()

        branch = request.form.get(
            "branch",
            ""
        ).strip()

        graduation_year = request.form.get(
            "graduation_year",
            ""
        ).strip()

        candidate_type = request.form.get(
            "candidate_type",
            "Fresher"
        ).strip()

        experience = request.form.get(
            "experience",
            ""
        ).strip()

        skills = request.form.get(
            "skills",
            ""
        ).strip()

        address = request.form.get(
            "address",
            ""
        ).strip()


        # ------------------------------------------
        # Required fields
        # ------------------------------------------

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

            "address": address

        }


        for field_name, field_value in required_fields.items():

            if not field_value:

                return render_template(
                    "apply.html",
                    error=(
                        field_name
                        .replace(
                            "_",
                            " "
                        )
                        .title()
                        + " is required."
                    )
                )


        connection = get_db_connection()


        # ------------------------------------------
        # Re-check verified account
        # ------------------------------------------

        user = connection.execute(
            """
            SELECT *
            FROM users
            WHERE id = ?
            """,
            (
                session["user_id"],
            )
        ).fetchone()


        if not user:

            connection.close()

            session.clear()

            return redirect(
                url_for(
                    "login"
                )
            )


        if user["email_verified"] != 1:

            connection.close()

            return render_template(
                "login.html",
                error=(
                    "Please verify your email before applying."
                )
            )


        # ------------------------------------------
        # One application per account
        # ------------------------------------------

        existing_application = connection.execute(
            """
            SELECT id
            FROM applications
            WHERE user_id = ?
            LIMIT 1
            """,
            (
                session["user_id"],
            )
        ).fetchone()


        if existing_application:

            connection.close()

            return redirect(
                url_for(
                    "my_application"
                )
            )


        # ------------------------------------------
        # Email already used
        # ------------------------------------------

        existing_email_application = connection.execute(
            """
            SELECT id
            FROM applications
            WHERE LOWER(email) = LOWER(?)
            LIMIT 1
            """,
            (
                email,
            )
        ).fetchone()


        if existing_email_application:

            connection.close()

            return render_template(
                "apply.html",
                error=(
                    "This email has already been used "
                    "to submit an application."
                )
            )


        # ------------------------------------------
        # Mobile already used
        # ------------------------------------------

        existing_mobile_application = connection.execute(
            """
            SELECT id
            FROM applications
            WHERE mobile = ?
            LIMIT 1
            """,
            (
                mobile,
            )
        ).fetchone()


        if existing_mobile_application:

            connection.close()

            return render_template(
                "apply.html",
                error=(
                    "This mobile number has already "
                    "been used to submit an application."
                )
            )


        # ------------------------------------------
        # Resume upload
        # ------------------------------------------

        resume = request.files.get(
            "resume"
        )

        resume_filename = None


        if resume and resume.filename:

            if not allowed_file(
                resume.filename
            ):

                connection.close()

                return render_template(
                    "apply.html",
                    error=(
                        "Invalid resume format. "
                        "Only PDF, DOC and DOCX files "
                        "are allowed."
                    )
                )


            original_filename = secure_filename(
                resume.filename
            )


            if not original_filename:

                connection.close()

                return render_template(
                    "apply.html",
                    error=(
                        "Invalid resume filename."
                    )
                )


            base_name, extension = os.path.splitext(
                original_filename
            )


            resume_filename = (
                f"{session['user_id']}_"
                f"{base_name}"
                f"{extension}"
            )


            resume_path = os.path.join(
                app.config["UPLOAD_FOLDER"],
                resume_filename
            )


            resume.save(
                resume_path
            )


        # ------------------------------------------
        # Save application
        # ------------------------------------------

        try:

            connection.execute(
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

                    resume_filename

                )

                VALUES (
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?
                )
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

                    resume_filename

                )
            )


            connection.commit()


            application_id = connection.execute(
                """
                SELECT last_insert_rowid()
                """
            ).fetchone()[0]


            application = connection.execute(
                """
                SELECT *
                FROM applications
                WHERE id = ?
                """,
                (
                    application_id,
                )
            ).fetchone()


        except Exception:

            connection.rollback()
            connection.close()


            if resume_filename:

                resume_path = os.path.join(
                    app.config["UPLOAD_FOLDER"],
                    resume_filename
                )


                if os.path.isfile(
                    resume_path
                ):

                    try:

                        os.remove(
                            resume_path
                        )

                    except OSError:

                        pass


            raise


        connection.close()


        # ------------------------------------------
        # Create application PDF
        # ------------------------------------------

        pdf_path = create_application_pdf(
            application
        )


        # ------------------------------------------
        # Send success email
        # ------------------------------------------

        email_sent = True


        try:

            send_application_success_email(
                application,
                pdf_path
            )


        except Exception as error:

            email_sent = False

            print(
                "Application success email error:",
                error
            )


        # ------------------------------------------
        # Email message
        # ------------------------------------------

        if email_sent:

            email_message = """
            <p>
                A confirmation email containing your
                application PDF has been sent to your
                registered email address.
            </p>
            """

        else:

            email_message = """
            <p style="color:#b00020;">
                Your application was saved successfully,
                but the confirmation email could not be sent.
                Please download the application PDF from
                your dashboard.
            </p>
            """


        # ------------------------------------------
        # Success page
        # ------------------------------------------

        safe_name = html.escape(
            str(application["name"])
        )


        safe_address = html.escape(
            str(WALKIN_OFFICE_ADDRESS)
        )

        safe_date = html.escape(
            str(WALKIN_DATE)
        )

        safe_time = html.escape(
            str(WALKIN_TIME)
        )

        safe_contact = html.escape(
            str(WALKIN_CONTACT)
        )


        return f"""
        <!DOCTYPE html>

        <html>

        <head>

            <title>
                Trekso | Application Submitted
            </title>

            <style>

                body {{
                    font-family:Arial,sans-serif;
                    background:#fff7f2;
                    text-align:center;
                    padding-top:70px;
                }}

                .box {{
                    background:white;
                    width:580px;
                    max-width:92%;
                    margin:auto;
                    padding:45px;
                    border-radius:14px;
                    box-shadow:
                        0 10px 35px
                        rgba(0,0,0,0.08);
                }}

                h2 {{
                    color:#ff4b00;
                }}

                p {{
                    color:#555;
                    line-height:1.6;
                }}

                .application-id {{
                    font-size:20px;
                    font-weight:bold;
                    color:#111;
                    margin:20px 0;
                }}

                .walkin {{
                    background:#fff7f2;
                    padding:20px;
                    border-radius:10px;
                    margin:25px 0;
                    text-align:left;
                }}

                .walkin h3 {{
                    color:#ff4b00;
                }}

                a {{
                    display:inline-block;
                    background:#ff4b00;
                    color:white;
                    text-decoration:none;
                    padding:12px 22px;
                    border-radius:7px;
                    font-weight:bold;
                }}

            </style>

        </head>

        <body>

            <div class="box">

                <h2>
                    Application Submitted Successfully!
                </h2>

                <p>
                    Dear {safe_name},
                </p>

                <p>
                    Your Trekso job application has been
                    successfully submitted.
                </p>

                <div class="application-id">
                    Application ID:
                    #{application["id"]}
                </div>

                {email_message}

                <div class="walkin">

                    <h3>
                        Walk-in Interview
                    </h3>

                    <p>
                        <strong>Office Address:</strong>
                        <br>
                        {safe_address}
                    </p>

                    <p>
                        <strong>Date:</strong>
                        <br>
                        {safe_date}
                    </p>

                    <p>
                        <strong>Time:</strong>
                        <br>
                        {safe_time}
                    </p>

                    <p>
                        <strong>Contact:</strong>
                        <br>
                        {safe_contact}
                    </p>

                    <p>
                        Please carry your application PDF
                        and resume when attending the interview.
                    </p>

                </div>

                <a href="/applicant-home">
                    Go to Dashboard
                </a>

            </div>

        </body>

        </html>
        """


    return render_template(
        "apply.html"
    )


# ==================================================
# VIEW MY APPLICATION
# ==================================================

@app.route(
    "/my-application"
)
def my_application():

    if "user_id" not in session:

        return redirect(
            url_for(
                "login"
            )
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT *
        FROM applications
        WHERE user_id = ?
        LIMIT 1
        """,
        (
            session["user_id"],
        )
    ).fetchone()


    connection.close()


    return render_template(
        "my_application.html",
        application=application
    )


# ==================================================
# APPLICANT DOWNLOAD RESUME
# ==================================================

@app.route(
    "/download-resume/<int:application_id>"
)
def download_resume(
    application_id
):

    if "user_id" not in session:

        return redirect(
            url_for(
                "login"
            )
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT *
        FROM applications
        WHERE id = ?
        AND user_id = ?
        """,
        (
            application_id,
            session["user_id"]
        )
    ).fetchone()


    connection.close()


    if not application:

        return (
            "Application not found or access denied.",
            404
        )


    resume_filename = application[
        "resume_filename"
    ]


    if not resume_filename:

        return (
            "No resume was uploaded.",
            404
        )


    resume_path = os.path.join(
        app.config["UPLOAD_FOLDER"],
        resume_filename
    )


    if not os.path.isfile(
        resume_path
    ):

        return (
            "Resume file not found.",
            404
        )


    return send_from_directory(
        app.config["UPLOAD_FOLDER"],
        resume_filename,
        as_attachment=True
    )


# ==================================================
# DOWNLOAD APPLICATION PDF
# ==================================================

@app.route(
    "/download-application"
)
def download_application():

    if "user_id" not in session:

        return redirect(
            url_for(
                "login"
            )
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT *
        FROM applications
        WHERE user_id = ?
        LIMIT 1
        """,
        (
            session["user_id"],
        )
    ).fetchone()


    connection.close()


    if not application:

        return (
            "No application found.",
            404
        )


    pdf_path = create_application_pdf(
        application
    )


    return send_file(
        pdf_path,
        as_attachment=True,
        download_name=(
            "Trekso_Application_"
            + str(application["id"])
            + ".pdf"
        )
    )


# ==================================================
# ADMIN LOGIN
# ==================================================

@app.route(
    "/admin-login",
    methods=["GET", "POST"]
)
def admin_login():

    ADMIN_USERNAME = os.environ.get(
        "ADMIN_USERNAME",
        "admin"
    )


    ADMIN_PASSWORD = os.environ.get(
        "ADMIN_PASSWORD",
        "trekso1245"
    )


    if request.method == "POST":

        username = request.form.get(
            "username",
            ""
        ).strip()

        password = request.form.get(
            "password",
            ""
        )


        if (
            username == ADMIN_USERNAME
            and password == ADMIN_PASSWORD
        ):

            session.clear()

            session.permanent = True

            session[
                "admin_logged_in"
            ] = True


            return redirect(
                url_for(
                    "admin_dashboard"
                )
            )


        return render_template(
            "admin_login.html",
            error=(
                "Invalid admin username "
                "or password."
            )
        )


    return render_template(
        "admin_login.html"
    )


# ==================================================
# ADMIN DASHBOARD
# ==================================================

@app.route(
    "/admin-dashboard"
)
def admin_dashboard():

    if not session.get(
        "admin_logged_in"
    ):

        return redirect(
            url_for(
                "admin_login"
            )
        )


    search = request.args.get(
        "search",
        ""
    ).strip()


    status_filter = request.args.get(
        "status",
        ""
    ).strip()


    connection = get_db_connection()


    query = """
        SELECT *
        FROM applications
        WHERE 1 = 1
    """


    parameters = []


    if search:

        query += """
            AND (
                name LIKE ?
                OR email LIKE ?
                OR mobile LIKE ?
                OR position LIKE ?
                OR qualification LIKE ?
                OR college_name LIKE ?
                OR university_name LIKE ?
                OR branch LIKE ?
            )
        """


        search_value = (
            f"%{search}%"
        )


        parameters.extend([

            search_value,

            search_value,

            search_value,

            search_value,

            search_value,

            search_value,

            search_value,

            search_value

        ])


    if status_filter:

        query += """
            AND status = ?
        """


        parameters.append(
            status_filter
        )


    query += """
        ORDER BY id DESC
    """


    applications = connection.execute(
        query,
        parameters
    ).fetchall()


    total_count = connection.execute(
        """
        SELECT COUNT(*)
        FROM applications
        """
    ).fetchone()[0]


    submitted_count = connection.execute(
        """
        SELECT COUNT(*)
        FROM applications
        WHERE status = 'Submitted'
        """
    ).fetchone()[0]


    selected_count = connection.execute(
        """
        SELECT COUNT(*)
        FROM applications
        WHERE status = 'Selected'
        """
    ).fetchone()[0]


    shortlisted_count = connection.execute(
        """
        SELECT COUNT(*)
        FROM applications
        WHERE status = 'Shortlisted'
        """
    ).fetchone()[0]


    connection.close()


    return render_template(
        "admin_dashboard.html",
        applications=applications,
        total_count=total_count,
        submitted_count=submitted_count,
        selected_count=selected_count,
        shortlisted_count=shortlisted_count,
        search=search,
        status_filter=status_filter
    )


# ==================================================
# ADMIN VIEW APPLICATION
# ==================================================

@app.route(
    "/admin-application/<int:application_id>"
)
def admin_view_application(
    application_id
):

    if not session.get(
        "admin_logged_in"
    ):

        return redirect(
            url_for(
                "admin_login"
            )
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT *
        FROM applications
        WHERE id = ?
        """,
        (
            application_id,
        )
    ).fetchone()


    connection.close()


    if not application:

        return (
            "Application not found.",
            404
        )


    return render_template(
        "admin_application.html",
        application=application
    )


# ==================================================
# ADMIN UPDATE STATUS
# ==================================================

@app.route(
    "/admin-update-status/<int:application_id>",
    methods=["POST"]
)
def update_application_status(
    application_id
):

    if not session.get(
        "admin_logged_in"
    ):

        return redirect(
            url_for(
                "admin_login"
            )
        )


    new_status = request.form.get(
        "status",
        ""
    ).strip()


    allowed_statuses = {

        "Submitted",

        "Under Review",

        "Shortlisted",

        "Selected",

        "Rejected"

    }


    if new_status not in allowed_statuses:

        return (
            "Invalid application status.",
            400
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT id
        FROM applications
        WHERE id = ?
        """,
        (
            application_id,
        )
    ).fetchone()


    if not application:

        connection.close()

        return (
            "Application not found.",
            404
        )


    connection.execute(
        """
        UPDATE applications
        SET status = ?
        WHERE id = ?
        """,
        (
            new_status,
            application_id
        )
    )


    connection.commit()
    connection.close()


    return redirect(
        url_for(
            "admin_view_application",
            application_id=application_id
        )
    )


# ==================================================
# ADMIN DELETE APPLICATION
# ==================================================

@app.route(
    "/admin-delete-application/<int:application_id>",
    methods=["POST"]
)
def admin_delete_application(
    application_id
):

    if not session.get(
        "admin_logged_in"
    ):

        return redirect(
            url_for(
                "admin_login"
            )
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT id, resume_filename
        FROM applications
        WHERE id = ?
        """,
        (
            application_id,
        )
    ).fetchone()


    if not application:

        connection.close()

        return (
            "Application not found.",
            404
        )


    resume_filename = (
        application["resume_filename"]
    )


    if resume_filename:

        resume_path = os.path.join(
            app.config["UPLOAD_FOLDER"],
            resume_filename
        )


        if os.path.isfile(
            resume_path
        ):

            try:

                os.remove(
                    resume_path
                )

            except OSError:

                pass


    connection.execute(
        """
        DELETE FROM applications
        WHERE id = ?
        """,
        (
            application_id,
        )
    )


    connection.commit()
    connection.close()


    return redirect(
        url_for(
            "admin_dashboard"
        )
    )


# ==================================================
# ADMIN DOWNLOAD RESUME
# ==================================================

@app.route(
    "/admin-download-resume/<int:application_id>"
)
def admin_download_resume(
    application_id
):

    if not session.get(
        "admin_logged_in"
    ):

        return redirect(
            url_for(
                "admin_login"
            )
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT *
        FROM applications
        WHERE id = ?
        """,
        (
            application_id,
        )
    ).fetchone()


    connection.close()


    if not application:

        return (
            "Application not found.",
            404
        )


    resume_filename = (
        application["resume_filename"]
    )


    if not resume_filename:

        return (
            "No resume was uploaded.",
            404
        )


    resume_path = os.path.join(
        app.config["UPLOAD_FOLDER"],
        resume_filename
    )


    if not os.path.isfile(
        resume_path
    ):

        return (
            "Resume file not found.",
            404
        )


    return send_from_directory(
        app.config["UPLOAD_FOLDER"],
        resume_filename,
        as_attachment=True
    )


# ==================================================
# ADMIN EXPORT EXCEL
# ==================================================

@app.route(
    "/admin-export-excel"
)
def admin_export_excel():

    if not session.get(
        "admin_logged_in"
    ):

        return redirect(
            url_for(
                "admin_login"
            )
        )


    connection = get_db_connection()


    applications = connection.execute(
        """
        SELECT *
        FROM applications
        ORDER BY id DESC
        """
    ).fetchall()


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

        "Submitted On"

    ]


    worksheet.append(
        headings
    )


    for application in applications:

        worksheet.append([

            application["id"],

            application["position"],

            application["name"],

            application["email"],

            application["mobile"],

            application["dob"],

            application["gender"],

            application["qualification"],

            application["college_name"],

            application["university_name"],

            application["branch"],

            application["graduation_year"],

            application["candidate_type"],

            application["experience"],

            application["skills"],

            application["address"],

            application["resume_filename"],

            application["status"],

            application["created_at"]

        ])


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

        "S": 22

    }


    for column, width in column_widths.items():

        worksheet.column_dimensions[
            column
        ].width = width


    worksheet.freeze_panes = "A2"


    excel_path = os.path.join(
        PDF_FOLDER,
        "applications.xlsx"
    )


    workbook.save(
        excel_path
    )


    return send_file(
        excel_path,
        as_attachment=True,
        download_name="applications.xlsx"
    )


# ==================================================
# ADMIN LOGOUT
# ==================================================

@app.route(
    "/admin-logout"
)
def admin_logout():

    session.clear()

    return redirect(
        url_for(
            "admin_login"
        )
    )


# ==================================================
# APPLICANT LOGOUT
# ==================================================

@app.route(
    "/logout"
)
def logout():

    session.clear()

    return redirect(
        url_for(
            "login"
        )
    )


# ==================================================
# FILE TOO LARGE
# ==================================================

@app.errorhandler(413)
def file_too_large(error):

    return """
    <h2 style="
        text-align:center;
        margin-top:100px;
        color:#ff4b00;
    ">
        Resume file is too large.
    </h2>

    <p style="
        text-align:center;
    ">
        Maximum allowed file size is 5 MB.
    </p>

    <p style="
        text-align:center;
    ">
        <a href="/apply">
            Go Back
        </a>
    </p>
    """, 413


# ==================================================
# RUN
# ==================================================

if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            5000
        )
    )


    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )