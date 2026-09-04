import os
import sqlite3


# ==================================================
# DATABASE CONFIGURATION
# ==================================================

DATABASE_TYPE = os.environ.get(
    "DATABASE_TYPE",
    "sqlite"
)


# ==================================================
# SQLITE CONNECTION
# ==================================================

def get_sqlite_connection():

    connection = sqlite3.connect(
        "hiring.db"
    )

    connection.row_factory = sqlite3.Row

    return connection


# ==================================================
# MAIN DATABASE CONNECTION
# ==================================================

def get_db_connection():

    # For now we continue using SQLite locally.
    # Later, DATABASE_TYPE can be changed to mysql
    # for the live Vercel deployment.

    if DATABASE_TYPE.lower() == "sqlite":

        return get_sqlite_connection()

    raise RuntimeError(
        "Unsupported DATABASE_TYPE: "
        + DATABASE_TYPE
    )


# ==================================================
# INITIALIZE DATABASE
# ==================================================

def init_db():

    connection = get_db_connection()


    # ------------------------------------------
    # Users table
    # ------------------------------------------

    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS users (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            name TEXT NOT NULL,

            email TEXT UNIQUE NOT NULL,

            mobile TEXT NOT NULL UNIQUE,

            password TEXT NOT NULL,

            email_verified INTEGER DEFAULT 0,

            mobile_verified INTEGER DEFAULT 0,

            email_otp TEXT,

            mobile_otp TEXT,

            otp_expiry TEXT

        )
        """
    )


    # ------------------------------------------
    # Applications table
    # ------------------------------------------

    connection.execute(
        """
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
            college_name TEXT NOT NULL,
university_name TEXT NOT NULL,

            branch TEXT NOT NULL,

            graduation_year TEXT NOT NULL,

            candidate_type TEXT NOT NULL
                DEFAULT 'Fresher',

            experience TEXT,

            skills TEXT NOT NULL,

            address TEXT NOT NULL,

            resume_filename TEXT,

            status TEXT DEFAULT 'Submitted',

            created_at TIMESTAMP
                DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (user_id)
                REFERENCES users(id)

        )
        """
    )


    connection.commit()

    connection.close()


# ==================================================
# DATABASE UPGRADE
# ==================================================


def upgrade_database():

    connection = get_db_connection()

    # ------------------------------------------
    # Check users table columns
    # ------------------------------------------

    user_columns = connection.execute(
        "PRAGMA table_info(users)"
    ).fetchall()

    user_column_names = [
        column["name"]
        for column in user_columns
    ]

    # Add email_verified if missing
    if "email_verified" not in user_column_names:
        connection.execute("""
            ALTER TABLE users
            ADD COLUMN email_verified INTEGER DEFAULT 0
        """)

    # Add mobile_verified if missing
    if "mobile_verified" not in user_column_names:
        connection.execute("""
            ALTER TABLE users
            ADD COLUMN mobile_verified INTEGER DEFAULT 0
        """)

    # Add email_otp if missing
    if "email_otp" not in user_column_names:
        connection.execute("""
            ALTER TABLE users
            ADD COLUMN email_otp TEXT
        """)

    # Add mobile_otp if missing
    if "mobile_otp" not in user_column_names:
        connection.execute("""
            ALTER TABLE users
            ADD COLUMN mobile_otp TEXT
        """)

    # Add otp_expiry if missing
    if "otp_expiry" not in user_column_names:
        connection.execute("""
            ALTER TABLE users
            ADD COLUMN otp_expiry TEXT
        """)


    # ------------------------------------------
    # Check applications table columns
    # ------------------------------------------

    application_columns = connection.execute(
        "PRAGMA table_info(applications)"
    ).fetchall()

    application_column_names = [
        column["name"]
        for column in application_columns
    ]

    # Add candidate_type if missing
    if "candidate_type" not in application_column_names:
        connection.execute("""
            ALTER TABLE applications
            ADD COLUMN candidate_type
            TEXT DEFAULT 'Fresher'
        """)

    # Add experience if missing
    if "experience" not in application_column_names:
        connection.execute("""
            ALTER TABLE applications
            ADD COLUMN experience TEXT
        """)

    # Add college_name if missing
    if "college_name" not in application_column_names:
        connection.execute("""
            ALTER TABLE applications
            ADD COLUMN college_name TEXT DEFAULT ''
        """)

    # Add university_name if missing
    if "university_name" not in application_column_names:
        connection.execute("""
            ALTER TABLE applications
            ADD COLUMN university_name TEXT DEFAULT ''
        """)


    # ------------------------------------------
    # Save changes
    # ------------------------------------------

    connection.commit()

    connection.close()
# ==================================================
# RUN DIRECTLY
# ==================================================

if __name__ == "__main__":

    init_db()

    upgrade_database()

    print(
        "Database initialized successfully!"
    )