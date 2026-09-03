import sqlite3

DATABASE = "hiring.db"


def get_db_connection():
    connection = sqlite3.connect(DATABASE)
    connection.row_factory = sqlite3.Row
    return connection


def init_db():

    connection = get_db_connection()

    # Users table
    connection.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            mobile TEXT NOT NULL,
            password TEXT NOT NULL,

            email_verified INTEGER DEFAULT 0,
            mobile_verified INTEGER DEFAULT 0,

            email_otp TEXT,
            mobile_otp TEXT,

            otp_expiry TEXT
        )
    """)

    # Applications table
    connection.execute("""
        CREATE TABLE IF NOT EXISTS applications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            position TEXT NOT NULL,
            name TEXT NOT NULL,
            email TEXT NOT NULL,
            mobile TEXT NOT NULL,
            dob TEXT NOT NULL,
            gender TEXT NOT NULL,
            qualification TEXT NOT NULL,
            branch TEXT NOT NULL,
            graduation_year TEXT NOT NULL,

            candidate_type TEXT NOT NULL DEFAULT 'Fresher',
            experience TEXT,

            skills TEXT NOT NULL,
            address TEXT NOT NULL,
            resume_filename TEXT,
            status TEXT DEFAULT 'Submitted',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (user_id) REFERENCES users(id)
        )
    """)

    connection.commit()
    connection.close()


def upgrade_database():

    connection = get_db_connection()

    # -----------------------------
    # Check USERS table
    # -----------------------------

    user_columns = connection.execute(
        "PRAGMA table_info(users)"
    ).fetchall()

    user_column_names = [
        column["name"]
        for column in user_columns
    ]

    # Email verification
    if "email_verified" not in user_column_names:
        connection.execute("""
            ALTER TABLE users
            ADD COLUMN email_verified INTEGER DEFAULT 0
        """)

    # Mobile verification
    if "mobile_verified" not in user_column_names:
        connection.execute("""
            ALTER TABLE users
            ADD COLUMN mobile_verified INTEGER DEFAULT 0
        """)

    # Email OTP
    if "email_otp" not in user_column_names:
        connection.execute("""
            ALTER TABLE users
            ADD COLUMN email_otp TEXT
        """)

    # Mobile OTP
    if "mobile_otp" not in user_column_names:
        connection.execute("""
            ALTER TABLE users
            ADD COLUMN mobile_otp TEXT
        """)

    # OTP expiry
    if "otp_expiry" not in user_column_names:
        connection.execute("""
            ALTER TABLE users
            ADD COLUMN otp_expiry TEXT
        """)


    # -----------------------------
    # Check APPLICATIONS table
    # -----------------------------

    application_columns = connection.execute(
        "PRAGMA table_info(applications)"
    ).fetchall()

    application_column_names = [
        column["name"]
        for column in application_columns
    ]

    # Candidate type
    if "candidate_type" not in application_column_names:
        connection.execute("""
            ALTER TABLE applications
            ADD COLUMN candidate_type TEXT DEFAULT 'Fresher'
        """)

    # Experience
    if "experience" not in application_column_names:
        connection.execute("""
            ALTER TABLE applications
            ADD COLUMN experience TEXT
        """)

    connection.commit()
    connection.close()


if __name__ == "__main__":

    init_db()

    upgrade_database()

    print("Database updated successfully!")