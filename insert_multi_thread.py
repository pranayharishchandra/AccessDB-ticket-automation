import pyodbc
import csv
from pathlib import Path

# ==============================
# CONFIG
# ==============================
BASE_DIR = Path(r"C:\Users\PranayHarishchandra\Desktop\s_workspace\test")
DELIMITER = "\t"
SKIP_DBS = {"database.accdb"}
BATCH_SIZE = 1000  # adjust if needed

# ==============================
# MAIN LOOP
# ==============================
for accdb_path in BASE_DIR.glob("*.accdb"):

    if accdb_path.name.lower() in SKIP_DBS:
        print(f"Skipping {accdb_path.name}")
        continue

    prefix = accdb_path.stem
    txt_files = sorted(BASE_DIR.glob(f"{prefix}*.txt"))

    if not txt_files:
        print(f"No TXT found for {prefix}, skipping")
        continue

    print(f"\nProcessing DB: {accdb_path.name}")
    print(f"TXT files: {[f.name for f in txt_files]}")

    # ==============================
    # CONNECT TO ACCESS
    # ==============================
    conn = pyodbc.connect(
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={accdb_path};"
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
        print("No user tables found, skipping DB")
        cursor.close()
        conn.close()
        continue

    target_table = sorted(tables)[0]
    print(f"Target table: {target_table}")

    # ==============================
    # GET COLUMN COUNT
    # ==============================
    cursor.execute(f"SELECT * FROM [{target_table}] WHERE 1=0")
    col_count = len(cursor.description)

    placeholders = ",".join("?" * col_count)
    insert_sql = f"INSERT INTO [{target_table}] VALUES ({placeholders})"

    inserted = 0
    skipped = 0
    batch = []

    # ==============================
    # PROCESS TXT FILES
    # ==============================
    for txt_file in txt_files:
        print(f"  Importing {txt_file.name}")

        with open(txt_file, newline="", encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=DELIMITER)

            for row in reader:

                if not row or all(not c.strip() for c in row):
                    skipped += 1
                    continue

                while len(row) > col_count and row[-1] == "":
                    row.pop()

                if len(row) != col_count:
                    skipped += 1
                    continue

                row = [c.strip() if c.strip() else None for c in row]

                batch.append(row)

                # 🔥 Insert when batch size reached
                if len(batch) >= BATCH_SIZE:
                    cursor.executemany(insert_sql, batch)
                    inserted += len(batch)
                    print(f"    Inserted batch of {len(batch)} rows")
                    batch.clear()

    # Insert remaining rows
    if batch:
        cursor.executemany(insert_sql, batch)
        inserted += len(batch)
        print(f"    Inserted final batch of {len(batch)} rows")

    conn.commit()
    cursor.close()
    conn.close()

    print(f"Inserted: {inserted}, Skipped: {skipped}")

print("\n✅ ALL DATABASES PROCESSED SUCCESSFULLY")
