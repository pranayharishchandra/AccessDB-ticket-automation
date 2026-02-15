import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import pyodbc
import csv

# ==============================
# CONFIG
# ==============================
DELIMITER = "\t"


# ==============================
# MAIN GUI CLASS
# ==============================
class DBToolApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Access DB Utility Tool")
        self.root.geometry("900x600")

        self.base_dir = Path.cwd()

        self.create_header()
        self.create_tabs()

    # ==============================
    # HEADER
    # ==============================
    def create_header(self):
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(header_frame, text="Folder Location:").pack(side="left")

        self.path_var = tk.StringVar(value=str(self.base_dir))
        self.path_entry = ttk.Entry(header_frame, textvariable=self.path_var, width=80)
        self.path_entry.pack(side="left", padx=5)

        ttk.Button(header_frame, text="Browse", command=self.browse_folder).pack(side="left")

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.base_dir = Path(folder)
            self.path_var.set(folder)
            self.refresh_all_lists()

    # ==============================
    # TABS
    # ==============================
    def create_tabs(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)

        # Empty DB tab
        self.empty_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.empty_tab, text="Empty DB")

        # Insert DB tab
        self.insert_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.insert_tab, text="Insert DB")

        self.create_tab_content(self.empty_tab, mode="empty")
        self.create_tab_content(self.insert_tab, mode="insert")

    # ==============================
    # TAB CONTENT
    # ==============================
    def create_tab_content(self, tab, mode):

        top_frame = ttk.Frame(tab)
        top_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(top_frame, text="Select .accdb files:").pack(anchor="w")

        list_frame = ttk.Frame(tab)
        list_frame.pack(fill="x", padx=10)

        canvas = tk.Canvas(list_frame, height=150)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)

        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="x", expand=True)
        scrollbar.pack(side="right", fill="y")

        tab.check_vars = {}
        tab.scrollable_frame = scrollable_frame

        # Run Button
        run_btn = ttk.Button(tab, text="RUN", command=lambda: self.run_action(mode, tab))
        run_btn.pack(pady=10)

        # Log area
        log_label = ttk.Label(tab, text="Logs:")
        log_label.pack(anchor="w", padx=10)

        log_text = tk.Text(tab, height=15)
        log_text.pack(fill="both", expand=True, padx=10, pady=5)

        tab.log_text = log_text

        self.populate_file_list(tab)

    # ==============================
    # FILE LIST POPULATION
    # ==============================
    def populate_file_list(self, tab):
        for widget in tab.scrollable_frame.winfo_children():
            widget.destroy()

        tab.check_vars.clear()

        for file in sorted(self.base_dir.glob("*.accdb")):
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(
                tab.scrollable_frame,
                text=file.name,
                variable=var
            )
            chk.pack(anchor="w")
            tab.check_vars[file] = var

    def refresh_all_lists(self):
        self.populate_file_list(self.empty_tab)
        self.populate_file_list(self.insert_tab)

    # ==============================
    # RUN ACTION
    # ==============================
    def run_action(self, mode, tab):

        selected_files = [file for file, var in tab.check_vars.items() if var.get()]

        if not selected_files:
            messagebox.showwarning("Warning", "No database selected!")
            return

        tab.log_text.delete("1.0", tk.END)

        for db_path in selected_files:
            if mode == "empty":
                self.clear_database(db_path, tab)
            else:
                self.insert_database(db_path, tab)

    # ==============================
    # EMPTY DATABASE
    # ==============================
    def clear_database(self, db_path, tab):

        self.log(tab, f"Clearing: {db_path.name}")

        try:
            conn = pyodbc.connect(
                r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                f"DBQ={db_path};"
            )
            cursor = conn.cursor()

            tables = [
                row.table_name
                for row in cursor.tables(tableType="TABLE")
                if not row.table_name.startswith("MSys")
            ]

            for table in tables:
                cursor.execute(f"DELETE FROM [{table}]")
                self.log(tab, f"  Cleared table: {table}")

            conn.commit()
            cursor.close()
            conn.close()

            self.log(tab, "  SUCCESS\n")

        except Exception as e:
            self.log(tab, f"  ERROR: {e}\n")

    # ==============================
    # INSERT DATABASE
    # ==============================
    def insert_database(self, db_path, tab):

        self.log(tab, f"Inserting into: {db_path.name}")

        try:
            conn = pyodbc.connect(
                r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                f"DBQ={db_path};"
            )
            cursor = conn.cursor()
            cursor.fast_executemany = True

            tables = [
                row.table_name
                for row in cursor.tables(tableType="TABLE")
                if not row.table_name.startswith("MSys")
            ]

            if not tables:
                self.log(tab, "  No user tables found\n")
                return

            tables.sort()
            target_table = tables[0]

            cursor.execute(f"SELECT * FROM [{target_table}] WHERE 1=0")
            col_count = len(cursor.description)

            placeholders = ",".join("?" * col_count)
            insert_sql = f"INSERT INTO [{target_table}] VALUES ({placeholders})"

            prefix = db_path.stem
            txt_files = sorted(self.base_dir.glob(f"{prefix}*.txt"))

            inserted = 0
            skipped = 0

            for txt_file in txt_files:
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

                        cursor.execute(insert_sql, row)
                        inserted += 1

            conn.commit()
            cursor.close()
            conn.close()

            self.log(tab, f"  Inserted: {inserted}, Skipped: {skipped}")
            self.log(tab, "  SUCCESS\n")

        except Exception as e:
            self.log(tab, f"  ERROR: {e}\n")

    # ==============================
    # LOG FUNCTION
    # ==============================
    def log(self, tab, message):
        tab.log_text.insert(tk.END, message + "\n")
        tab.log_text.see(tk.END)


# ==============================
# START APP
# ==============================
if __name__ == "__main__":
    root = tk.Tk()
    app = DBToolApp(root)
    root.mainloop()
