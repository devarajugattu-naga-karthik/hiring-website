import os
import psycopg2
from psycopg2.extras import RealDictCursor


# ==================================================
# DATABASE CONFIGURATION
# ==================================================

DATABASE_URL = os.environ.get("DATABASE_URL")


# ==================================================
# DATABASE CONNECTION WRAPPER
# ==================================================

class DatabaseConnection:

    def __init__(self):

        if not DATABASE_URL:

            raise RuntimeError(
                "DATABASE_URL environment variable is not configured."
            )

        self.connection = psycopg2.connect(
            DATABASE_URL
        )


    def execute(
        self,
        query,
        parameters=None
    ):

        cursor = self.connection.cursor(
            cursor_factory=RealDictCursor
        )

        cursor.execute(
            query,
            parameters or ()
        )

        return DatabaseCursor(
            self.connection,
            cursor
        )


    def commit(self):

        self.connection.commit()


    def rollback(self):

        self.connection.rollback()


    def close(self):

        self.connection.close()


# ==================================================
# CURSOR WRAPPER
# ==================================================

class DatabaseCursor:

    def __init__(
        self,
        connection,
        cursor
    ):

        self.connection = connection
        self.cursor = cursor


    def fetchone(self):

        return self.cursor.fetchone()


    def fetchall(self):

        return self.cursor.fetchall()


    def __getitem__(self, key):

        return self.cursor.fetchone()[key]


    def __iter__(self):

        return iter(
            self.cursor.fetchall()
        )


    def close(self):

        self.cursor.close()


# ==================================================
# MAIN DATABASE CONNECTION
# ==================================================

def get_db_connection():

    return DatabaseConnection()


# ==================================================
# INITIALIZE DATABASE
# ==================================================

def init_db():

    connection = get_db_connection()


    # ==================================================
    # USERS
    # ==================================================

    result = connection.execute(
        """
        CREATE TABLE IF NOT EXISTS users (

            id BIGSERIAL PRIMARY KEY,

            name TEXT NOT NULL,

            email TEXT UNIQUE NOT NULL,

            mobile TEXT NOT NULL UNIQUE,

            password TEXT NOT NULL,

            email_verified INTEGER DEFAULT 1,

            mobile_verified INTEGER DEFAULT 0,

            email_otp TEXT,

            mobile_otp TEXT,

            otp_expiry TIMESTAMP,

            otp_attempts INTEGER DEFAULT 0,

            created_at TIMESTAMP
                DEFAULT CURRENT_TIMESTAMP

        )
        """
    )

    result.close()


    # ==================================================
    # PENDING REGISTRATIONS
    # ==================================================

    result = connection.execute(
        """
        CREATE TABLE IF NOT EXISTS pending_registrations (

            id BIGSERIAL PRIMARY KEY,

            name TEXT NOT NULL,

            email TEXT UNIQUE NOT NULL,

            mobile TEXT NOT NULL,

            password_hash TEXT NOT NULL,

            otp TEXT NOT NULL,

            otp_expiry TIMESTAMP NOT NULL,

            otp_attempts INTEGER DEFAULT 0,

            created_at TIMESTAMP
                DEFAULT CURRENT_TIMESTAMP

        )
        """
    )

    result.close()


    # ==================================================
    # APPLICATIONS
    # ==================================================

    result = connection.execute(
        """
        CREATE TABLE IF NOT EXISTS applications (

            id BIGSERIAL PRIMARY KEY,

            user_id BIGINT NOT NULL,

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

            CONSTRAINT fk_applications_user
                FOREIGN KEY (user_id)
                REFERENCES users(id)

        )
        """
    )

    result.close()


    connection.commit()

    connection.close()


# ==================================================
# DATABASE UPGRADE
# ==================================================

def upgrade_database():

    connection = get_db_connection()


    # ==================================================
    # USERS UPGRADES
    # ==================================================

    statements = [

        """
        ALTER TABLE users
        ADD COLUMN IF NOT EXISTS email_verified
        INTEGER DEFAULT 1
        """,

        """
        ALTER TABLE users
        ADD COLUMN IF NOT EXISTS mobile_verified
        INTEGER DEFAULT 0
        """,

        """
        ALTER TABLE users
        ADD COLUMN IF NOT EXISTS email_otp
        TEXT
        """,

        """
        ALTER TABLE users
        ADD COLUMN IF NOT EXISTS mobile_otp
        TEXT
        """,

        """
        ALTER TABLE users
        ADD COLUMN IF NOT EXISTS otp_expiry
        TIMESTAMP
        """,

        """
        ALTER TABLE users
        ADD COLUMN IF NOT EXISTS otp_attempts
        INTEGER DEFAULT 0
        """,

        """
        ALTER TABLE users
        ADD COLUMN IF NOT EXISTS created_at
        TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        """

    ]


    for statement in statements:

        result = connection.execute(
            statement
        )

        result.close()


    # ==================================================
    # APPLICATION UPGRADES
    # ==================================================

    application_statements = [

        """
        ALTER TABLE applications
        ADD COLUMN IF NOT EXISTS candidate_type
        TEXT DEFAULT 'Fresher'
        """,

        """
        ALTER TABLE applications
        ADD COLUMN IF NOT EXISTS experience
        TEXT
        """,

        """
        ALTER TABLE applications
        ADD COLUMN IF NOT EXISTS college_name
        TEXT DEFAULT ''
        """,

        """
        ALTER TABLE applications
        ADD COLUMN IF NOT EXISTS university_name
        TEXT DEFAULT ''
        """

    ]


    for statement in application_statements:

        result = connection.execute(
            statement
        )

        result.close()


    # ==================================================
    # INDEXES
    # ==================================================

    result = connection.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_applications_user_id
        ON applications(user_id)
        """
    )

    result.close()


    result = connection.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_applications_email
        ON applications(email)
        """
    )

    result.close()


    result = connection.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_applications_mobile
        ON applications(mobile)
        """
    )

    result.close()


    connection.commit()

    connection.close()


# ==================================================
# TEST DATABASE
# ==================================================

if __name__ == "__main__":

    init_db()

    upgrade_database()

    print(
        "PostgreSQL database initialized successfully!"
    )