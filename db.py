import sqlite3
import os

DB_PATH = os.path.join("database", "mapping.db")


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    return conn

def create_table_dynamic(columns):
    """
    columns = ["NO", "NO RM", "NAMA PASIEN", ...]
    membuat tabel mapping otomatis berdasarkan header Excel
    """
    conn = get_connection()
    cursor = conn.cursor()

    cols = ", ".join([f"{col} TEXT" for col in columns])

    sql = f"""
        CREATE TABLE IF NOT EXISTS mapping (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            {cols}
        )
    """
    cursor.execute(sql)
    conn.commit()
    conn.close()

def insert_dynamic(row_dict):
    conn = get_connection()
    cursor = conn.cursor()

    keys = ", ".join(row_dict.keys())
    values_qm = ", ".join(["?"] * len(row_dict))
    values = list(row_dict.values())

    sql = f"INSERT INTO mapping ({keys}) VALUES ({values_qm})"

    cursor.execute(sql, values)
    conn.commit()
    conn.close()
