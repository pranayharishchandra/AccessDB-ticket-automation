import pyodbc
import csv

# ==============================
# CONFIG
# ==============================

ACCESS_DB = r"C:\\Users\\PranayHarishchandra\\Desktop\\s_workspace\\LQUA.accdb"
TXT_FILE = r"C:\\Users\\PranayHarishchandra\\Desktop\\s_workspace\\LQUA.txt"
DELIMITER = "\t"

# ==============================
# CONNECT TO ACCESS
# ==============================
conn = pyodbc.connect(
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    rf"DBQ={ACCESS_DB};"
)
cursor = conn.cursor()
cursor.fast_executemany = True

# ==============================
# FIND FIRST USER TABLE
# ==============================
tables = [
    row.table_name
    for row in cursor.tables(tableType="TABLE")
    if not row.table_name.startswith("MSys")
]

if not tables:
    raise RuntimeError("No user tables found in database")

tables.sort()
target_table = tables[0]

print(f"Target table: {target_table}")

# ==============================
# GET COLUMN COUNT FROM TABLE
# ==============================
cursor.execute(f"SELECT * FROM [{target_table}] WHERE 1=0")
table_column_count = len(cursor.description)

print(f"Table column count: {table_column_count}")

# ==============================
# READ TXT AND INSERT DATA
# ==============================
inserted = 0
skipped = 0

with open(TXT_FILE, newline="", encoding="utf-8") as f:
    reader = csv.reader(f, delimiter=DELIMITER)

    for line_no, row in enumerate(reader, start=1):

        # Skip fully empty rows
        if not row or all(not col.strip() for col in row):
            skipped += 1
            continue

        # Remove SAP trailing empty column (caused by trailing TAB)
        while len(row) > table_column_count and row[-1] == "":
            row.pop()

        # Validate column count
        if len(row) != table_column_count:
            print(f"Skipping line {line_no}: column mismatch ({len(row)})")
            skipped += 1
            continue

        # Convert empty strings to NULL (Access requirement)
        row = [col.strip() if col.strip() != "" else None for col in row]

        placeholders = ",".join("?" * table_column_count)
        sql = f"INSERT INTO [{target_table}] VALUES ({placeholders})"

        cursor.execute(sql, row)
        inserted += 1

# ==============================
# COMMIT & CLEANUP
# ==============================
conn.commit()
cursor.close()
conn.close()

print(f"Inserted rows: {inserted}")
print(f"Skipped rows: {skipped}")
print("Import completed successfully.")
