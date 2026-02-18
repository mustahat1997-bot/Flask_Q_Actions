import os
import sqlite3
import pandas as pd
import re

# =========================
# PATHS & SETTINGS
# =========================
#ONEDRIVE_FOLDER = r"C:\Users\lenovo\Earthlink Telecommunications\Rauf Samim - Q 40 and 30 sites"
TARGET_SHEET = "All Repeaters & Affiliates"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "Q_Actions.db")
EXCEL_PATH = os.path.join(BASE_DIR, "Updated changes queues Repeaters Data V148  Final).xlsx")

#DB_PATH = r"C:\Users\lenovo\Desktop\Python\Q_Actions.db"

TABLE_NAME = "repeater_actions"


# =========================
# NORMALIZE TEXT (🔥 مهم)
# =========================
def normalize_text(text):
    return re.sub(r"\s+", "", str(text)).lower().strip()


# =========================
# GET LATEST EXCEL FILE
# =========================
def get_latest_excel_file(folder_path):
    excel_files = [
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith(".xlsx")
    ]

    if not excel_files:
        raise FileNotFoundError("No Excel files found")

    return max(excel_files, key=os.path.getmtime)


# =========================
# EXCEL -> DATAFRAME
# =========================
def excel_to_dataframe():
    excel_file = get_latest_excel_file(EXCEL_PATH)

    df = pd.read_excel(excel_file, sheet_name=TARGET_SHEET)

    df.columns = df.columns.str.strip()

    df = df.rename(columns={
        "RepeaterName / Affiliates Name": "name",
        "Site code": "site_code",
        "Q Action": "q_action",
        "Repeater Action": "repeater_action"
    })

    required_columns = ["name", "site_code", "q_action", "repeater_action"]
    df = df[required_columns]

    # تنظيف القيم
    df["name"] = df["name"].astype(str).str.strip()
    df["site_code"] = df["site_code"].astype(str).str.strip()
    df["q_action"] = df["q_action"].fillna("No Action")
    df["repeater_action"] = df["repeater_action"].fillna("No Action")

    if df.empty:
        raise ValueError("DataFrame is empty")

    return df, os.path.basename(excel_file)


# =========================
# SAVE TO SQLITE
# =========================
def save_to_sqlite(df):
    print("🗄 Using DB:", DB_PATH)

    conn = sqlite3.connect(DB_PATH)
    df.to_sql(TABLE_NAME, conn, if_exists="replace", index=False)
    conn.close()

    print("✅ Database updated successfully")


# =========================
# DB CONNECTION
# =========================
def connect_db():
    return sqlite3.connect(DB_PATH)


# =========================
# GET ACTIONS FROM DB
# البحث بدون فراغات + بدون فرق كابيتال
# =========================
def get_actions_from_db(cursor, value):

    normalized_value = normalize_text(value)

    print("🔎 Searching DB for:", normalized_value)

    cursor.execute("""
        SELECT q_action, repeater_action
        FROM repeater_actions
        WHERE REPLACE(LOWER(name), ' ', '') = ?
        OR REPLACE(LOWER(site_code), ' ', '') = ?
        LIMIT 1
    """, (normalized_value, normalized_value))

    row = cursor.fetchone()

    return row if row else ("No Action", "No Action")


# =========================
# GET Q ACTION BY SITE CODE
# =========================
def get_q_action_by_site_code(cursor, site_code):

    normalized_value = normalize_text(site_code)

    print("🔎 Searching DB by Site Code:", normalized_value)

    cursor.execute("""
        SELECT q_action
        FROM repeater_actions
        WHERE REPLACE(LOWER(site_code), ' ', '') = ?
        LIMIT 1
    """, (normalized_value,))

    row = cursor.fetchone()

    return row[0] if row else "No Action"


# =========================
# GET ACTION TYPE
# =========================
def get_action_type(rule):
    if rule.startswith("(+"):
        return "add"
    if rule.startswith("(-"):
        return "decrease"
    return "add"


# =========================
# MAIN (Run manually if needed)
# =========================
def main():
    print("🔍 Searching latest Excel file...")
    df, file_name = excel_to_dataframe()
    print(f"💾 Updating DB from: {file_name}")
    save_to_sqlite(df)


if __name__ == "__main__":
    main()
