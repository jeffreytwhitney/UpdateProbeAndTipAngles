import os
import sys
from pathlib import Path
from typing import List, Dict, Any

import pymssql


def _iter_env_file_candidates() -> list[Path]:
    """Return candidate .env locations in priority order."""
    candidates: list[Path] = []

    if getattr(sys, 'frozen', False):
        candidates.append(Path(sys.executable).resolve().parent / '.env')

    candidates.append(Path.cwd() / '.env')
    candidates.append(Path(__file__).resolve().parent / '.env')

    deduped: list[Path] = []
    seen: set[Path] = set()
    for candidate in candidates:
        resolved = candidate.resolve(strict=False)
        if resolved not in seen:
            seen.add(resolved)
            deduped.append(resolved)

    return deduped


def _load_dotenv_file() -> None:
    """Load key=value pairs from a local .env file without overriding existing env vars."""
    env_path = next((path for path in _iter_env_file_candidates() if path.exists()), None)
    if env_path is None:
        return

    for raw_line in env_path.read_text(encoding='utf-8').splitlines():
        line = raw_line.strip()
        if not line or line.startswith('#') or '=' not in line:
            continue

        key, value = line.split('=', 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            os.environ.setdefault(key, value)


def _get_required_env(name: str) -> str:
    value = os.getenv(name)
    if not value or not value.strip():
        raise RuntimeError(
            f"Missing required environment variable '{name}'. "
            f"Copy .env.example to .env and set your local values."
        )
    return value.strip()


_load_dotenv_file()


class DatabaseConnection:
    DB_SERVER_ENV = 'FBCR_DB_SERVER'
    DB_USER_ENV = 'FBCR_DB_USER'
    DB_PASSWORD_ENV = 'FBCR_DB_PASSWORD'
    DB_NAME_ENV = 'FBCR_DB_NAME'

    def __init__(self, is_scalar: bool = False):
        self.conn = None
        self.cursor = None
        self.is_scalar = is_scalar

    def __enter__(self):
        server = _get_required_env(self.DB_SERVER_ENV)
        user = _get_required_env(self.DB_USER_ENV)
        password = _get_required_env(self.DB_PASSWORD_ENV)
        database = _get_required_env(self.DB_NAME_ENV)

        if self.is_scalar:
            self.conn = pymssql.connect(
                server=server,
                user=user,
                password=password,
                database=database,
                as_dict=False)
        else:
            self.conn = pymssql.connect(
                server=server,
                user=user,
                password=password,
                database=database,
                as_dict=True)

        self.cursor = self.conn.cursor()

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.cursor:
            self.cursor.close()
        if self.conn:
            self.conn.close()

    def fetch_records(self, sql: str) -> List[Dict[str, Any]]:
        self.cursor.execute(sql)
        return self.cursor.fetchall()

    def fetch_scalar(self, sql: str) -> Any:
        self.cursor.execute(sql)
        return self.cursor.fetchone()[0]

    def execute_scalar_statement(self, sql: str) -> Any:
        """Execute a mutating statement that returns one scalar value and commit."""
        self.cursor.execute(sql)
        result = self.cursor.fetchone()[0]
        self.conn.commit()
        return result

    def execute_statement(self, sql: str) -> None:
        self.cursor.execute(sql)
        self.conn.commit()

    def execute_stored_procedure(self, procedure_name: str, params: tuple[Any, ...] | None = None) -> None:
        self.cursor.callproc(procedure_name, params or ())
        self.conn.commit()


class WindowsAuthDatabaseConnection:
    """Database connection class that uses Windows authentication."""
    
    DB_SERVER_ENV = 'FBCR_DB_SERVER'
    
    def __init__(self, database_name: str, is_scalar: bool = False):
        """
        Initialize the connection with Windows authentication.
        
        Args:
            database_name: The name of the database to connect to
            is_scalar: If True, returns results as tuples; if False, returns as dictionaries
        """
        self.database_name = database_name
        self.conn = None
        self.cursor = None
        self.is_scalar = is_scalar

    def __enter__(self):
        # Windows authentication is used when user and password are not provided
        server = _get_required_env(self.DB_SERVER_ENV)
        if self.is_scalar:
            self.conn = pymssql.connect(
                server=server,
                database=self.database_name,
                as_dict=False)
        else:
            self.conn = pymssql.connect(
                server=server,
                database=self.database_name,
                as_dict=True)

        self.cursor = self.conn.cursor()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.cursor:
            self.cursor.close()
        if self.conn:
            self.conn.close()

    def fetch_records(self, sql: str) -> List[Dict[str, Any]]:
        """Execute a query and return all records."""
        self.cursor.execute(sql)
        return self.cursor.fetchall()

    def fetch_scalar(self, sql: str) -> Any:
        """Execute a query and return a single scalar value."""
        self.cursor.execute(sql)
        return self.cursor.fetchone()[0]

    def execute_scalar_statement(self, sql: str) -> Any:
        """Execute a mutating statement that returns one scalar value and commit."""
        self.cursor.execute(sql)
        result = self.cursor.fetchone()[0]
        self.conn.commit()
        return result

    def execute_statement(self, sql: str) -> None:
        """Execute a SQL statement (INSERT, UPDATE, DELETE, etc.)."""
        self.cursor.execute(sql)
        self.conn.commit()

    def execute_stored_procedure(self, procedure_name: str, params: tuple[Any, ...] | None = None) -> None:
        """Execute a stored procedure using Windows authentication."""
        self.cursor.callproc(procedure_name, params or ())
        self.conn.commit()


def get_sql_recordset(sql: str) -> List[Dict[str, Any]]:
    with DatabaseConnection(False) as db:
        return db.fetch_records(sql)


def get_sql_scalar(sql: str) -> Any:
    with DatabaseConnection(True) as db:
        return db.fetch_scalar(sql)


def execute_sql_scalar_statement(sql: str) -> Any:
    """Execute a mutating statement that returns one scalar value."""
    with DatabaseConnection(True) as db:
        return db.execute_scalar_statement(sql)


def execute_sql_statement(sql: str) -> None:
    with DatabaseConnection(False) as db:
        db.execute_statement(sql)


def execute_stored_procedure(procedure_name: str, params: tuple[Any, ...] | None = None) -> None:
    with DatabaseConnection(False) as db:
        db.execute_stored_procedure(procedure_name, params)


def get_sql_recordset_windows_auth(database_name: str, sql: str) -> List[Dict[str, Any]]:
    """Execute a query with Windows authentication and return all records."""
    with WindowsAuthDatabaseConnection(database_name, is_scalar=False) as db:
        return db.fetch_records(sql)


def get_sql_scalar_windows_auth(database_name: str, sql: str) -> Any:
    """Execute a query with Windows authentication and return a single scalar value."""
    with WindowsAuthDatabaseConnection(database_name, is_scalar=True) as db:
        return db.fetch_scalar(sql)


def execute_sql_scalar_statement_windows_auth(database_name: str, sql: str) -> Any:
    """Execute a mutating statement that returns one scalar value with Windows authentication."""
    with WindowsAuthDatabaseConnection(database_name, is_scalar=True) as db:
        return db.execute_scalar_statement(sql)


def execute_sql_statement_windows_auth(database_name: str, sql: str) -> None:
    """Execute a SQL statement with Windows authentication."""
    with WindowsAuthDatabaseConnection(database_name, is_scalar=False) as db:
        db.execute_statement(sql)


def execute_stored_procedure_windows_auth(
        database_name: str,
        procedure_name: str,
        params: tuple[Any, ...] | None = None) -> None:
    """Execute a stored procedure with Windows authentication."""
    with WindowsAuthDatabaseConnection(database_name, is_scalar=False) as db:
        db.execute_stored_procedure(procedure_name, params)

