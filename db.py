import os
import psycopg2
from psycopg2.extras import RealDictCursor
from dotenv import load_dotenv

load_dotenv()

def db():
    conn = psycopg2.connect(
        host=os.getenv("PGHOST", "localhost"),
        port=os.getenv("PGPORT", "5432"),
        dbname=os.getenv("PGDATABASE", "postgres"),
        user=os.getenv("PGUSER", "postgres"),
        password=os.getenv("PGPASSWORD", ""),
        sslmode=os.getenv("PGSSLMODE", "require")
    )
    # Устанавливаем схему после подключения
    with conn.cursor() as cur:
        cur.execute(f"SET search_path TO {os.getenv('PGSCHEMA', 'public')}")
    return conn

def rows(sql, params=()):
    """Выполнить SELECT и вернуть список словарей"""
    with db() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(sql, params)
            return cur.fetchall()

def execute(sql, params=()):
    """Выполнить INSERT/UPDATE/DELETE"""
    with db() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)

def ref_data():
    return {
        "users": rows("SELECT id AS user_id, full_name, role FROM users ORDER BY full_name"),
        "animals": rows("""
            SELECT id AS animal_id, 
                   COALESCE(nickname, tag_number, 'Животное #' || id) AS label,
                   tag_number, nickname, group_id, status
            FROM animals ORDER BY label
        """),
        "groups": rows('SELECT id AS group_id, name FROM "groups" ORDER BY name')
    }
