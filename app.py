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

# Session lifetime = 24 hours
app.permanent_session_lifetime = 60 * 60 * 24


# ==================================================
# UPLOAD SETTINGS
# ==================================================

UPLOAD_FOLDER = "uploads"

ALLOWED_EXTENSIONS = {
    "pdf",
    "doc",
    "docx"
}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# Maximum resume size = 5 MB
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024

os.makedirs(
    UPLOAD_FOLDER,
    exist_ok=True
)


# ==================================================
# DATABASE
# ==================================================

init_db()
upgrade_database()


# ==================================================
# HELPER FUNCTION
# ==================================================

def allowed_file(filename):

    return (
        "." in filename
        and filename.rsplit(
            ".",
            1
        )[1].lower() in ALLOWED_EXTENSIONS
    )


# ==================================================
# HOME PAGE
# ==================================================

@app.route("/")
def home():

    return render_template(
        "home.html"
    )


# ==================================================
# APPLICANT REGISTRATION
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
        # Basic validation
        # ------------------------------------------

        if not name or not email or not mobile or not password:

            return render_template(
                "register.html",
                error="All fields are required."
            )


        connection = get_db_connection()


        # ------------------------------------------
        # Check duplicate email
        # ------------------------------------------

        existing_email = connection.execute(
            """
            SELECT id
            FROM users
            WHERE LOWER(email) = LOWER(?)
            """,
            (email,)
        ).fetchone()


        if existing_email:

            connection.close()

            return render_template(
                "register.html",
                error="Email already registered."
            )


        # ------------------------------------------
        # Check duplicate mobile
        # ------------------------------------------

        existing_mobile = connection.execute(
            """
            SELECT id
            FROM users
            WHERE mobile = ?
            """,
            (mobile,)
        ).fetchone()


        if existing_mobile:

            connection.close()

            return render_template(
                "register.html",
                error="Mobile number already registered."
            )


        # ------------------------------------------
        # Hash password
        # ------------------------------------------

        hashed_password = generate_password_hash(
            password
        )


        # ------------------------------------------
        # Create user
        # ------------------------------------------

        connection.execute(
            """
            INSERT INTO users (
                name,
                email,
                mobile,
                password
            )
            VALUES (?, ?, ?, ?)
            """,
            (
                name,
                email,
                mobile,
                hashed_password
            )
        )


        connection.commit()
        connection.close()


        return redirect(
            url_for("login")
        )


    return render_template(
        "register.html"
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
            (email,)
        ).fetchone()


        connection.close()


        # ------------------------------------------
        # Successful applicant login
        # ------------------------------------------

        if user and check_password_hash(
            user["password"],
            password
        ):

            session.clear()

            session.permanent = True

            session["user_id"] = user["id"]
            session["user_name"] = user["name"]
            session["user_email"] = user["email"]
            session["user_mobile"] = user["mobile"]


            return redirect(
                url_for("applicant_home")
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

@app.route("/applicant-home")
def applicant_home():

    if "user_id" not in session:

        return redirect(
            url_for("login")
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT *
        FROM applications
        WHERE user_id = ?
        LIMIT 1
        """,
        (session["user_id"],)
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
            url_for("login")
        )


    # ------------------------------------------
    # Check existing application
    # ------------------------------------------

    connection = get_db_connection()


    existing_application = connection.execute(
        """
        SELECT id, position
        FROM applications
        WHERE user_id = ?
        LIMIT 1
        """,
        (session["user_id"],)
    ).fetchone()


    connection.close()


    # ------------------------------------------
    # Prevent second application
    # ------------------------------------------

    if existing_application:

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
                    box-shadow:
                        0 10px 35px
                        rgba(0,0,0,0.08);
                }}

                h2 {{
                    color: #ff4b00;
                    margin-bottom: 20px;
                }}

                p {{
                    color: #555;
                    margin-bottom: 15px;
                }}

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

                <h2>
                    Application Already Submitted
                </h2>

                <p>
                    You have already applied for:
                </p>

                <p>
                    <strong>
                        {existing_application["position"]}
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


    # ------------------------------------------
    # Process application
    # ------------------------------------------

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
        # Validate required fields
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
                    error=f"{field_name.replace('_', ' ').title()} is required."
                )


        connection = get_db_connection()


        # ------------------------------------------
        # Double-check one application rule
        # ------------------------------------------

        existing_application = connection.execute(
            """
            SELECT id
            FROM applications
            WHERE user_id = ?
            LIMIT 1
            """,
            (session["user_id"],)
        ).fetchone()


        if existing_application:

            connection.close()

            return redirect(
                url_for("my_application")
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
                        "Only PDF, DOC and DOCX "
                        "files are allowed."
                    )
                )


            original_filename = secure_filename(
                resume.filename
            )


            if not original_filename:

                connection.close()

                return render_template(
                    "apply.html",
                    error="Invalid resume filename."
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
        connection.close()


        # ------------------------------------------
        # Success page
        # ------------------------------------------

        return """
        <!DOCTYPE html>

        <html>

        <head>

            <title>Teckso | Application Submitted</title>

            <style>

                body {
                    font-family: Arial, sans-serif;
                    background: #fff7f2;
                    text-align: center;
                    padding-top: 100px;
                }

                .box {
                    background: white;
                    width: 520px;
                    max-width: 90%;
                    margin: auto;
                    padding: 45px;
                    border-radius: 14px;
                    box-shadow:
                        0 10px 35px
                        rgba(0,0,0,0.08);
                }

                h2 {
                    color: #ff4b00;
                    margin-bottom: 20px;
                }

                p {
                    color: #555;
                    margin-bottom: 20px;
                }

                a {
                    display: inline-block;
                    background: #ff4b00;
                    color: white;
                    text-decoration: none;
                    padding: 12px 22px;
                    border-radius: 7px;
                    font-weight: bold;
                }

            </style>

        </head>

        <body>

            <div class="box">

                <h2>
                    Application Submitted Successfully!
                </h2>

                <p>
                    Your Teckso job application has been submitted successfully.
                </p>

                <p>
                    One applicant can submit only one application.
                </p>

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

@app.route("/my-application")
def my_application():

    if "user_id" not in session:

        return redirect(
            url_for("login")
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT *
        FROM applications
        WHERE user_id = ?
        LIMIT 1
        """,
        (session["user_id"],)
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
def download_resume(application_id):

    if "user_id" not in session:

        return redirect(
            url_for("login")
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

@app.route("/download-application")
def download_application():

    if "user_id" not in session:

        return redirect(
            url_for("login")
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT *
        FROM applications
        WHERE user_id = ?
        LIMIT 1
        """,
        (session["user_id"],)
    ).fetchone()


    connection.close()


    if not application:

        return (
            "No application found.",
            404
        )


    pdf_folder = "generated_pdfs"

    os.makedirs(
        pdf_folder,
        exist_ok=True
    )


    pdf_filename = (
        "application_"
        + str(application["id"])
        + ".pdf"
    )


    pdf_path = os.path.join(
        pdf_folder,
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
        textColor=colors.HexColor("#111111"),
        spaceAfter=10
    )


    subtitle_style = ParagraphStyle(
        "SubtitleStyle",
        parent=styles["Heading2"],
        alignment=TA_CENTER,
        textColor=colors.HexColor("#ff4b00"),
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
            "TECKSO",
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
        Spacer(1, 8)
    )


    story.append(
        Paragraph(
            "<b>Application Status:</b> "
            + str(application["status"]),
            normal_style
        )
    )


    story.append(
        Spacer(1, 20)
    )


    data = [

        ["Field", "Details"],

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
            str(application["experience"] or "N/A")
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
            str(application["resume_filename"] or "Not uploaded")
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
                colors.HexColor("#ff4b00")
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
        Spacer(1, 25)
    )


    story.append(
        Paragraph(
            "This document was generated electronically by Teckso Careers.",
            footer_style
        )
    )


    document.build(
        story
    )


    return send_file(
        pdf_path,
        as_attachment=True,
        download_name=pdf_filename
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
                url_for("admin_dashboard")
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

@app.route("/admin-dashboard")
def admin_dashboard():

    if not session.get(
        "admin_logged_in"
    ):

        return redirect(
            url_for("admin_login")
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


        search_value = f"%{search}%"


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


    # ------------------------------------------
    # Statistics
    # ------------------------------------------

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
            url_for("admin_login")
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT *
        FROM applications
        WHERE id = ?
        """,
        (application_id,)
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
# ADMIN UPDATE APPLICATION STATUS
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
            url_for("admin_login")
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
        (application_id,)
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
            url_for("admin_login")
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT id, resume_filename
        FROM applications
        WHERE id = ?
        """,
        (application_id,)
    ).fetchone()


    if not application:

        connection.close()

        return (
            "Application not found.",
            404
        )


    # ------------------------------------------
    # Delete resume file
    # ------------------------------------------

    resume_filename = application[
        "resume_filename"
    ]


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


    # ------------------------------------------
    # Delete database record
    # ------------------------------------------

    connection.execute(
        """
        DELETE FROM applications
        WHERE id = ?
        """,
        (application_id,)
    )


    connection.commit()
    connection.close()


    return redirect(
        url_for("admin_dashboard")
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
            url_for("admin_login")
        )


    connection = get_db_connection()


    application = connection.execute(
        """
        SELECT *
        FROM applications
        WHERE id = ?
        """,
        (application_id,)
    ).fetchone()


    connection.close()


    if not application:

        return (
            "Application not found.",
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
# ADMIN EXPORT TO EXCEL
# ==================================================

@app.route("/admin-export-excel")
def admin_export_excel():

    if not session.get(
        "admin_logged_in"
    ):

        return redirect(
            url_for("admin_login")
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


    # ------------------------------------------
    # Create workbook
    # ------------------------------------------

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


    # ------------------------------------------
    # Add applications
    # ------------------------------------------

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


    # ------------------------------------------
    # Column widths
    # ------------------------------------------

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


    # ------------------------------------------
    # Save Excel file
    # ------------------------------------------

    excel_folder = "generated_pdfs"


    os.makedirs(
        excel_folder,
        exist_ok=True
    )


    excel_path = os.path.join(
        excel_folder,
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

@app.route("/admin-logout")
def admin_logout():

    session.clear()

    return redirect(
        url_for("admin_login")
    )


# ==================================================
# APPLICANT LOGOUT
# ==================================================

@app.route("/logout")
def logout():

    session.clear()

    return redirect(
        url_for("login")
    )


# ==================================================
# ERROR: FILE TOO LARGE
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
# RUN APPLICATION
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