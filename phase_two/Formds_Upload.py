# Uploading the data from excel file to MySQL database 
# made by : Eslam Ayman 
# Date : 8 / 10 / 2025
# 
# Inputs to configure (edit these values before running):
# EXCEL_FILE : path to the Excel file (xlsx)
# HOST       : MySQL host
# USER       : MySQL username
# PASSWORD   : MySQL password
# DB_NAME    : target database name
# TABLE_NAME : target table name (will be created or replaced)
#
# Example:
# EXCEL_FILE = "Downloads/formds_2025-10-07.xlsx"
# HOST = "localhost"
# USER = "root"
# PASSWORD = "if you have put it or keep it empty"
# DB_NAME = "formds_db"
# TABLE_NAME = "scraped_data"

import os
import re
import sys
import logging
import pandas as pd
import mysql.connector
from mysql.connector import Error

# ====== CONFIGURE HERE ======
EXCEL_FILE = ""
HOST = "localhost"
USER = ""
PASSWORD = ""
DB_NAME = ""
TABLE_NAME = ""
# ============================

EXPECTED_COLUMNS = [
    "Order_num",
    "Company",
    "Reported_funds",
    "Incremental_cash",
    "Filing_date",
    "New_or_amended",
    "Company_address",
    "Company_all_contact",
]

logging.basicConfig(level=logging.INFO, format="%(message)s")


def sanitize_column_name(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r'\s+', '_', name)
    name = re.sub(r'[^\w]', '', name)
    return name or "col"


def normalize(name: str) -> str:
    return re.sub(r'[^a-z0-9_]', '', str(name).lower())


def read_excel(file_path: str) -> pd.DataFrame:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Excel file not found: {file_path}")
    return pd.read_excel(file_path, engine="openpyxl")


def connect_mysql(host, user, password, db_name=None):
    conn = mysql.connector.connect(host=host, user=user, password=password)
    cur = conn.cursor()
    if db_name:
        cur.execute(f"CREATE DATABASE IF NOT EXISTS `{db_name}`")
        conn.database = db_name
    return conn, cur


def table_exists(cursor, table_name):
    cursor.execute(
        "SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = DATABASE() AND table_name = %s",
        (table_name,),
    )
    return cursor.fetchone()[0] == 1


def get_table_columns(cursor, table_name):
    cursor.execute(f"SHOW COLUMNS FROM `{table_name}`")
    rows = cursor.fetchall()
    return [r[0] for r in rows]


def drop_table(cursor, table_name):
    cursor.execute(f"DROP TABLE IF EXISTS `{table_name}`")


def create_table_with_schema(cursor, table_name):
    sql = f"""
    CREATE TABLE `{table_name}` (
      `Order_num` INT(11) NOT NULL AUTO_INCREMENT,
      `Company` VARCHAR(500),
      `Reported_funds` VARCHAR(100),
      `Incremental_cash` VARCHAR(100),
      `Filing_date` DATE,
      `New_or_amended` VARCHAR(20),
      `Company_address` VARCHAR(500),
      `Company_all_contact` VARCHAR(2000),
      PRIMARY KEY (`Order_num`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """
    cursor.execute(sql)


def truncate_table(cursor, table_name):
    cursor.execute(f"TRUNCATE TABLE `{table_name}`")


def parse_date(value):
    try:
        if pd.isna(value):
            return None
        dt = pd.to_datetime(value, errors='coerce')
        if pd.isna(dt):
            return None
        return dt.date().isoformat()
    except Exception:
        return None


def find_col_index_by_candidates(cols_normalized, candidates):
    for cand in candidates:
        n = normalize(cand)
        if n in cols_normalized:
            return cols_normalized.index(n)
    return None


def insert_data(cursor, table_name, df):
    cols_sql = "`Company`,`Reported_funds`,`Incremental_cash`,`Filing_date`,`New_or_amended`,`Company_address`,`Company_all_contact`"
    placeholders = ", ".join(["%s"] * 7)
    insert_sql = f"INSERT INTO `{table_name}` ({cols_sql}) VALUES ({placeholders})"

    cols = [str(c) for c in df.columns.tolist()]
    cols_normalized = [normalize(c) for c in cols]

    idx_first = 0
    idx_second = 1

    idx_reported = find_col_index_by_candidates(cols_normalized, [
        "Reported_Funding", "ReportedFunding", "Reported_funds", "Reported_Funds", "ReportedFunding", "reported_funds"
    ])
    idx_incremental = find_col_index_by_candidates(cols_normalized, [
        "Incremental_Cash", "IncrementalCash", "Incremental_cash", "incremental_cash"
    ])
    idx_date = find_col_index_by_candidates(cols_normalized, [
        "Date", "Filing_date", "FilingDate", "date", "filing_date"
    ])
    idx_new = find_col_index_by_candidates(cols_normalized, [
        "New_or_Amended", "New_or_amended", "new_or_amended", "new_or_amend"
    ])

    total = 0
    for _, row in df.iterrows():
        try:
            a = "" if pd.isna(row.iloc[idx_first]) else str(row.iloc[idx_first]).strip()
        except Exception:
            a = ""
        try:
            b = "" if pd.isna(row.iloc[idx_second]) else str(row.iloc[idx_second]).strip()
        except Exception:
            b = ""
        company = (a + " " + b).strip() if (a or b) else None

        reported = None
        if idx_reported is not None:
            try:
                val = row.iloc[idx_reported]
                reported = None if pd.isna(val) else str(val).strip()
            except Exception:
                reported = None

        incremental = None
        if idx_incremental is not None:
            try:
                val = row.iloc[idx_incremental]
                incremental = None if pd.isna(val) else str(val).strip()
            except Exception:
                incremental = None

        filing_date = None
        if idx_date is not None:
            filing_date = parse_date(row.iloc[idx_date])

        new_or = None
        if idx_new is not None:
            try:
                val = row.iloc[idx_new]
                new_or = None if pd.isna(val) else str(val).strip()
                if new_or and len(new_or) > 20:
                    new_or = new_or[:20]
            except Exception:
                new_or = None

        data = [company, reported, incremental, filing_date, new_or, None, None]
        cursor.execute(insert_sql, data)
        total += 1

    return total


def main():
    conn = None
    cur = None
    try:
        logging.info("1) Reading Excel file...")
        df = read_excel(EXCEL_FILE)

        if df.shape[1] < 2:
            logging.error("Error: Excel file must contain at least two columns.")
            sys.exit(1)

        original_cols = list(df.columns)
        df.columns = [sanitize_column_name(c) for c in original_cols]
        logging.info(f"Columns after sanitization: {df.columns.tolist()}")

        logging.info("2) Connecting to MySQL...")
        conn, cur = connect_mysql(HOST, USER, PASSWORD, DB_NAME)

        needs_create = True
        if table_exists(cur, TABLE_NAME):
            existing_cols = get_table_columns(cur, TABLE_NAME)
            logging.info("Table exists. Checking schema...")
            if set(existing_cols) == set(EXPECTED_COLUMNS):
                needs_create = False
                logging.info("Table schema matches. Will truncate and insert new data.")
            else:
                logging.info("Table schema differs. Dropping and recreating the table with expected schema.")
                drop_table(cur, TABLE_NAME)
                conn.commit()
                needs_create = True

        if needs_create:
            logging.info("Creating table...")
            create_table_with_schema(cur, TABLE_NAME)
            conn.commit()
            logging.info("Table created.")

        logging.info("Truncating table...")
        truncate_table(cur, TABLE_NAME)
        conn.commit()

        logging.info("Inserting data from Excel into MySQL...")
        inserted = insert_data(cur, TABLE_NAME, df)
        conn.commit()
        logging.info(f"Inserted {inserted} rows into `{TABLE_NAME}` in database `{DB_NAME}`")

    except FileNotFoundError as e:
        logging.error(f"Error: {e}")
    except Error as e:
        logging.error(f"MySQL Error: {e}")
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
    finally:
        try:
            if cur:
                cur.close()
            if conn:
                conn.close()
        except:
            pass


if __name__ == "__main__":
    main()
