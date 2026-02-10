import json
import logging
import sys
import win32com.client


# ----------------------------
# CONFIG & LOGGING
# ----------------------------

def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler("run.log"),
            logging.StreamHandler(sys.stdout)
        ]
    )


def load_config(path="tables.json"):
    with open(path, "r") as f:
        return json.load(f)


# ----------------------------
# ACCESS HELPERS
# ----------------------------

def open_access_db(db_path):
    logging.info(f"Opening DB: {db_path}")
    access = win32com.client.Dispatch("Access.Application")
    access.OpenCurrentDatabase(db_path)
    access.Visible = False
    return access


def close_access_db(access):
    logging.info("Closing DB")
    access.CloseCurrentDatabase()
    access.Quit()


# ----------------------------
# CORE LOGIC
# ----------------------------

def clear_table(access, table_name):
    logging.info(f"Clearing table: {table_name}")

    sql_delete = f"DELETE FROM {table_name};"
    # access.DoCmd.RunSQL(sql_delete)
    access.CurrentDb().Execute(sql_delete) # Using Execute is more efficient and does not prompt for confirmation.

    # Validation
    rs = access.CurrentDb().OpenRecordset(
        f"SELECT COUNT(*) AS CNT FROM {table_name};"
    )
    count = rs.Fields("CNT").Value
    rs.Close()

    if count != 0:
        raise RuntimeError(f"Table {table_name} not empty after delete")


def get_valid_saved_imports(access, table_name):
    """
    Returns list of saved import names to run
    """
    logging.info("Detecting saved imports")

    valid_imports = []

    for imp in access.CurrentProject().AllDataAccessPages:
        # not used, placeholder
        pass

    # Correct way: use DoCmd.RunSavedImportExport
    # Since Access does not expose a clean collection,
    # we rely on naming convention from config or rules.

    # SKELETON: hard rule-based placeholder
    all_imports = access.CurrentDb().Containers("Scripts").Documents

    for imp in all_imports:
        name = imp.Name

        if table_name in name and not name.endswith("_ALL"):
            valid_imports.append(name)

    valid_imports.sort()
    return valid_imports


def run_saved_imports(access, import_names):
    for imp_name in import_names:
        logging.info(f"Running saved import: {imp_name}")
        access.DoCmd.RunSavedImportExport(imp_name)


def validate_load(access, table_name):
    rs = access.CurrentDb().OpenRecordset(
        f"SELECT COUNT(*) AS CNT FROM {table_name};"
    )
    count = rs.Fields("CNT").Value
    rs.Close()

    if count <= 0:
        raise RuntimeError(f"No data loaded into {table_name}")

    logging.info(f"{table_name} loaded successfully ({count} rows)")


# ----------------------------
# MAIN DRIVER
# ----------------------------

def main():
    setup_logging()
    config = load_config()

    for table_name, meta in config.items():
        logging.info(f"===== START TABLE {table_name} =====")

        access = None
        try:
            access = open_access_db(meta["db_path"])

            clear_table(access, meta["table_name"])

            imports = get_valid_saved_imports(access, table_name)

            if not imports:
                raise RuntimeError("No saved imports found")

            run_saved_imports(access, imports)

            validate_load(access, meta["table_name"])

            logging.info(f"===== SUCCESS {table_name} =====")

        except Exception as e:
            logging.error(f"FAILED for {table_name}: {e}")
            if access:
                close_access_db(access)
            sys.exit(1)

        close_access_db(access)

    logging.info("ALL TABLES PROCESSED SUCCESSFULLY")


if __name__ == "__main__":
    main()
