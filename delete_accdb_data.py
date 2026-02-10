import sys
import os
from pathlib import Path
import pyodbc


def find_accdb_files(directory: Path):
    """Return sorted list of .accdb files in directory."""
    return sorted([p for p in directory.iterdir() if p.is_file() and p.suffix.lower() == ".accdb"])


def clear_database(db_path: Path):
    """Connect to Access DB at db_path and delete all rows from user tables."""
    print(f"\nClearing database: {db_path.name}")
    conn = None
    cursor = None
    try:
        conn_str = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            + "DBQ="
            + str(db_path)
            + ";"
        )
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()

        tables = []
        for row in cursor.tables(tableType="TABLE"):
            table_name = row.table_name
            # Skip system tables
            if not table_name.startswith("MSys"):
                tables.append(table_name)

        if not tables:
            print("  No user tables found.")
            return

        for table in tables:
            print(f"  Deleting data from table: {table}")
            cursor.execute(f"DELETE FROM [{table}]")

        conn.commit()
        print("  Done: all user tables cleared.")

    except Exception as e:
        print(f"  Error while clearing {db_path.name}: {e}")
    finally:
        try:
            if cursor is not None:
                cursor.close()
        except Exception:
            pass
        try:
            if conn is not None:
                conn.close()
        except Exception:
            pass


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
