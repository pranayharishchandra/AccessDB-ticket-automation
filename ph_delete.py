import pyodbc
from pathlib import Path


def clear_database(db_path: Path):
    print(f"\nClearing database: {db_path.name}")

    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        + f"DBQ={db_path};"
    )

    conn = pyodbc.connect(conn_str, autocommit=False)
    cursor = conn.cursor()

    try:
        tables = [
            row.table_name
            for row in cursor.tables(tableType="TABLE")
            if not row.table_name.startswith("MSys")
        ]

        if not tables:
            print("  No user tables found.")
            return

        for table in tables:
            print(f"  Deleting data from: {table}")
            cursor.execute(f"DELETE FROM [{table}]")

        conn.commit()
        print("  Database cleared successfully.")

    except Exception as e:
        conn.rollback()
        raise RuntimeError(f"Clear failed for {db_path.name}: {e}")

    finally:
        cursor.close()
        conn.close()



def main():
    cwd = Path.cwd()
    accdb_files = find_accdb_files(cwd)

    if not accdb_files:
        print(f"No .accdb files found in the current directory: {cwd}")
        return

    print("Select a database to clear (enter the number):\n")
    for i, p in enumerate(accdb_files, start=1):
        print(f"{i}. {p.name}")
    all_index = len(accdb_files) + 1
    print(f"{all_index}. All files")

    try:
        choice = input("\nEnter number and press Enter: ").strip()
    except (EOFError, KeyboardInterrupt):
        print("\nInput cancelled.")
        return

    if not choice.isdigit():
        print("Invalid selection: not a number.")
        return

    choice_num = int(choice)
    if choice_num < 1 or choice_num > all_index:
        print("Selection out of range.")
        return

    targets = accdb_files if choice_num == all_index else [accdb_files[choice_num - 1]]

    # Optional confirmation
    print("\nYou selected:")
    for t in targets:
        print(f" - {t.name}")
    try:
        confirm = input("Type YES to confirm and clear the selected database(s): ").strip()
    except (EOFError, KeyboardInterrupt):
        print("\nConfirmation cancelled.")
        return

    if confirm != "YES":
        print("Confirmation not received. Aborting.")
        return

    for db in targets:
        clear_database(db)


if __name__ == "__main__":
    main()
