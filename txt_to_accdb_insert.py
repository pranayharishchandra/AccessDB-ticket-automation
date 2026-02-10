import pyodbc
import csv

conn = pyodbc.connect(
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=C:\Users\PranayHarishchandra\Desktop\s_workspace\LQUA.accdb;"
)
cursor = conn.cursor()

txt_file = r"C:\Users\PranayHarishchandra\Desktop\s_workspace\LQUA.TXT"


tables = []

for row in cursor.tables(tableType="TABLE"):
    table_name = row.table_name
    if not table_name.startswith("MSys"):  # skip system tables
        tables.append(table_name)

if not tables:
    raise Exception("No user tables found in database")

tables.sort()  # make order deterministic
tar_table_name = tables[0]


print(f"Inserting data from {txt_file} into table: {tar_table_name}")

cols = [c.column_name for c in cursor.columns(table=tar_table_name)]

head_counter = 1
for i in cols:
    print(f"  {head_counter}: {i}")
    head_counter += 1


print("----------------------------------------------------")

with open(txt_file, newline="", encoding="utf-8") as f:
    reader = csv.reader(f, delimiter="\t")

    for row in reader:
        # Skip completely empty rows
        if not row or all(not col.strip() for col in row):
            continue

        # Remove SAP trailing empty column
        if row[-1] == "":
            row = row[:-1]

        # Convert empty strings to NULL
        row = [col.strip() if col.strip() != "" else None for col in row]

        placeholders = ",".join(["?"] * len(row))
        sql = f"INSERT INTO [{tar_table_name}] VALUES ({placeholders})"

        cursor.execute(sql, row)

conn.commit()
cursor.close()
conn.close()
