import pandas as pd
import re
from db import create_table_dynamic, insert_dynamic

def normalize_column(col):
    col = col.lower()
    col = re.sub(r"[\s\-.]+", "_", col)
    col = re.sub(r"[^a-z0-9_]", "", col)
    return col

def fix_value(val):
    if pd.isna(val):
        return ""
    if isinstance(val, pd.Timestamp):
        return val.strftime("%Y-%m-%d")
    return str(val)

def save_mapping_to_db(file_path):
    # baca excel
    df = pd.read_excel(file_path)

    # normalisasi nama header
    df.columns = [normalize_column(c) for c in df.columns]

    # buat tabel berdasarkan header excel
    create_table_dynamic(df.columns)

    # insert setiap row secara dinamis
    for _, row in df.iterrows():
        row_dict = {col: fix_value(row[col]) for col in df.columns}
        insert_dynamic(row_dict)

    return len(df)
