import sys
import os
import datetime
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from db_executor import (
    get_available_drivers,
    test_connection,
    run_filtered_query,
    process_excel_pure_dmvic_recon,
    pyodbc,
    pd,
    openpyxl
)

# ==============================================================================
# HARDCODED DATABASE CONNECTION CREDENTIALS
# DO NOT MODIFY OR OVERWRITE USER CREDENTIALS BELOW
# ==============================================================================
SERVER = "172.26.0.21"           # Database Server Host / IP
DATABASE = "SIRIUS_MW"           # Database Name
USERNAME = "rptmw"               # SQL Server Username
PASSWORD = "Mwrpt$2020"          # SQL Server Password
AUTH_TYPE = "SQL Server Authentication"  # Authentication Type
DRIVER = "ODBC Driver 17 for SQL Server"  # ODBC Driver
# ==============================================================================

APP_NAME = "PURE-DMVIC RECON"
APP_VERSION = "Version 1.1"

class SQLRunnerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_NAME} - {APP_VERSION}")
        self.root.geometry("1100x750")
        self.root.minsize(900, 620)

        # Set app style
        self.style = ttk.Style()
        if "vista" in self.style.theme_names():
            self.style.theme_use("vista")
        elif "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        # Auto-resolve best ODBC driver
        available_drivers = get_available_drivers()
        self.resolved_driver = DRIVER if DRIVER in available_drivers else (available_drivers[0] if available_drivers else DRIVER)

        # Data state
        self.uploaded_excel_path = ""
        self.processed_df = None

        # Build UI layout
        self.create_widgets()

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)

        # Auto-test system connectivity silently on launch
        self.root.after(300, self.async_test_connection)

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="12")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # -------------------------------------------------------------
        # 1. Header & Version Banner
        # -------------------------------------------------------------
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 6))

        title_label = ttk.Label(
            header_frame,
            text=f"{APP_NAME}  |  {APP_VERSION}",
            font=("Segoe UI", 14, "bold")
        )
        title_label.pack(side=tk.LEFT)

        self.status_badge = ttk.Label(
            header_frame,
            text="[ Checking System Status... ]",
            font=("Segoe UI", 9, "bold"),
            foreground="#DCDCAA"
        )
        self.status_badge.pack(side=tk.RIGHT)

        # User Instruction Note Banner
        instruction_frame = ttk.Frame(main_frame, padding="6")
        instruction_frame.pack(fill=tk.X, pady=(0, 10))

        instruction_lbl = ttk.Label(
            instruction_frame,
            text="💡 Instructions: Select the year range for your DMVIC report, then upload the report from DMVIC and click 'Extract Report'.",
            font=("Segoe UI", 9, "italic"),
            foreground="#005A9E",
            wraplength=1040
        )
        instruction_lbl.pack(anchor=tk.W)

        # -------------------------------------------------------------
        # 2. Changes Log Section
        # -------------------------------------------------------------
        changes_frame = ttk.LabelFrame(main_frame, text="Changes Log (v1.1)", padding="8")
        changes_frame.pack(fill=tk.X, pady=(0, 10))

        changes_text = ScrolledText(changes_frame, height=8, wrap=tk.WORD, font=("Segoe UI", 9), bg="#F9F9F9")
        changes_text.pack(fill=tk.BOTH, expand=True)

        # Populate with static change log
        change_log = (
            "- Added year-range selector UI (Start/End Year).\n"
            "- Implemented document_ref filter for premium‑register query.\n"
            "- Integrated new SQL query with DebitNote period handling.\n"
            "- Generated CombinedReport worksheet with match flag.\n"
            "- Highlight rows where a registration has multiple DebitNotes.\n"
            "- Renamed application to PURE‑DMVIC RECON, Version 1.1.\n"
            "- Added this Changes Log section summarising all modifications.\n"
        )
        changes_text.insert(tk.END, change_log)
        changes_text.configure(state='disabled')

        # -------------------------------------------------------------
        # 2. Step 1: Year Range Filter (Start Year & End Year)
        # -------------------------------------------------------------
        filter_group = ttk.LabelFrame(main_frame, text=" Step 1: Select Report Year Range ", padding="10")
        filter_group.pack(fill=tk.X, pady=(0, 10))

        current_year = datetime.date.today().year
        year_list = [str(y) for y in range(2015, 2031)]

        # Start Year
        ttk.Label(filter_group, text="Start Year:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=(5, 5), pady=4, sticky=tk.W)
        self.start_year_var = tk.StringVar(value=str(current_year))
        self.start_year_combo = ttk.Combobox(filter_group, textvariable=self.start_year_var, values=year_list, width=8, font=("Segoe UI", 10), state="readonly")
        self.start_year_combo.grid(row=0, column=1, padx=(0, 20), pady=4, sticky=tk.W)

        # End Year
        ttk.Label(filter_group, text="End Year:", font=("Segoe UI", 9, "bold")).grid(row=0, column=2, padx=(5, 5), pady=4, sticky=tk.W)
        self.end_year_var = tk.StringVar(value=str(current_year))
        self.end_year_combo = ttk.Combobox(filter_group, textvariable=self.end_year_var, values=year_list, width=8, font=("Segoe UI", 10), state="readonly")
        self.end_year_combo.grid(row=0, column=3, padx=(0, 20), pady=4, sticky=tk.W)

        # Quick Presets
        preset_frame = ttk.Frame(filter_group)
        preset_frame.grid(row=0, column=4, padx=5, pady=4, sticky=tk.E)
        filter_group.columnconfigure(4, weight=1)

        ttk.Button(preset_frame, text=f"This Year ({current_year})", command=self.set_preset_current_year).pack(side=tk.LEFT, padx=2)
        ttk.Button(preset_frame, text="2025 - 2026", command=self.set_preset_2025_2026).pack(side=tk.LEFT, padx=2)
        ttk.Button(preset_frame, text="2020 - 2026", command=self.set_preset_2020_2026).pack(side=tk.LEFT, padx=2)

        # -------------------------------------------------------------
        # 3. Step 2: Upload DMVIC Report File
        # -------------------------------------------------------------
        excel_group = ttk.LabelFrame(main_frame, text=" Step 2: Upload Report from DMVIC ", padding="10")
        excel_group.pack(fill=tk.X, pady=(0, 10))
        excel_group.columnconfigure(0, weight=1)

        excel_input_frame = ttk.Frame(excel_group)
        excel_input_frame.pack(fill=tk.X)
        excel_input_frame.columnconfigure(0, weight=1)

        self.excel_path_var = tk.StringVar()
        self.excel_entry = ttk.Entry(excel_input_frame, textvariable=self.excel_path_var, font=("Segoe UI", 10))
        self.excel_entry.grid(row=0, column=0, sticky=tk.EW, padx=(0, 8))

        self.browse_excel_btn = ttk.Button(
            excel_input_frame,
            text="📁 Select DMVIC Report...",
            command=self.browse_excel_file
        )
        self.browse_excel_btn.grid(row=0, column=1)

        # -------------------------------------------------------------
        # 4. Step 3: Action Buttons (Extract Report)
        # -------------------------------------------------------------
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        self.process_btn = tk.Button(
            action_frame,
            text="▶  Extract Report",
            font=("Segoe UI", 11, "bold"),
            bg="#0078D4",
            fg="white",
            activebackground="#005A9E",
            activeforeground="white",
            relief="raised",
            bd=2,
            padx=25,
            pady=8,
            cursor="hand2",
            command=self.async_process_excel
        )
        self.process_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.save_excel_btn = tk.Button(
            action_frame,
            text="💾  Save Extracted Report...",
            font=("Segoe UI", 10, "bold"),
            bg="#107C41",
            fg="white",
            activebackground="#0B5C30",
            activeforeground="white",
            relief="raised",
            bd=2,
            padx=15,
            pady=6,
            cursor="hand2",
            state="disabled",
            command=self.save_processed_excel
        )
        self.save_excel_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.clear_btn = ttk.Button(action_frame, text="Clear View", command=self.clear_results)
        self.clear_btn.pack(side=tk.RIGHT)

        self.progressbar = ttk.Progressbar(main_frame, mode="indeterminate")

        # -------------------------------------------------------------
        # 5. Notebook (Preview & Logs)
        # -------------------------------------------------------------
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Tab 1: Preview Table
        self.table_tab = ttk.Frame(self.notebook, padding="4")
        self.notebook.add(self.table_tab, text=" Extracted Report Preview ")

        tree_frame = ttk.Frame(self.table_tab)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree_scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        self.tree_scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)

        self.tree = ttk.Treeview(
            tree_frame,
            yscrollcommand=self.tree_scroll_y.set,
            xscrollcommand=self.tree_scroll_x.set,
            selectmode="extended"
        )
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Tab 2: Logs
        self.log_tab = ttk.Frame(self.notebook, padding="4")
        self.notebook.add(self.log_tab, text=" Activity Logs ")

        self.log_text = ScrolledText(
            self.log_tab,
            wrap=tk.WORD,
            font=("Consolas", 10),
            bg="#1E1E1E",
            fg="#D4D4D4",
            insertbackground="white"
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.log_text.tag_config("timestamp", foreground="#808080")
        self.log_text.tag_config("info", foreground="#569CD6")
        self.log_text.tag_config("success", foreground="#4EC9B0")
        self.log_text.tag_config("error", foreground="#F44747")
        self.log_text.tag_config("warning", foreground="#DCDCAA")

    # -------------------------------------------------------------
    # Helper Methods
    # -------------------------------------------------------------
    def browse_excel_file(self):
        filename = filedialog.askopenfilename(
            title="Select Report from DMVIC",
            filetypes=[("Excel Files (*.xlsx;*.xls)", "*.xlsx;*.xls"), ("CSV Files (*.csv)", "*.csv"), ("All Files", "*.*")]
        )
        if filename:
            self.excel_path_var.set(filename)
            self.uploaded_excel_path = filename
            self.log(f"Loaded DMVIC report file: {filename}", "info")


    def set_preset_current_year(self):
        current_year = str(datetime.date.today().year)
        self.start_year_var.set(current_year)
        self.end_year_var.set(current_year)

    def set_preset_2025_2026(self):
        self.start_year_var.set("2025")
        self.end_year_var.set("2026")

    def set_preset_2020_2026(self):
        self.start_year_var.set("2020")
        self.end_year_var.set("2026")

    def log(self, message: str, tag: str = "info"):
        now_str = datetime.datetime.now().strftime("[%H:%M:%S] ")
        self.log_text.insert(tk.END, now_str, "timestamp")
        self.log_text.insert(tk.END, f"{message}\n", tag)
        self.log_text.see(tk.END)

    def clear_results(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = ()
        self.processed_df = None
        self.save_excel_btn.config(state="disabled")
        self.log_text.delete("1.0", tk.END)

    def set_running_state(self, is_running: bool):
        if is_running:
            self.process_btn.config(state="disabled", bg="#666666")
            self.browse_excel_btn.config(state="disabled")
            self.progressbar.pack(fill=tk.X, pady=(0, 5), before=self.notebook)
            self.progressbar.start(10)
        else:
            self.process_btn.config(state="normal", bg="#0078D4")
            self.browse_excel_btn.config(state="normal")
            self.progressbar.stop()
            self.progressbar.pack_forget()

    def update_treeview(self, df):
        if df is None:
            return
        self.tree.delete(*self.tree.get_children())
        columns = list(df.columns)
        self.tree["columns"] = columns
        self.tree["show"] = "headings"

        # Configure column headings and widths
        for col in columns:
            self.tree.heading(col, text=str(col))
            width = 140 if "DEBITNOTE" in str(col).upper() else 110
            self.tree.column(col, minwidth=80, width=width, anchor=tk.W)

        # Determine RegNo to DebitNote mapping for highlighting
        reg_col = None
        debit_col = None
        for idx, col_name in enumerate(columns):
            if col_name.lower() == "regno":
                reg_col = idx
            if col_name.lower() == "debitnote":
                debit_col = idx
        # Fallback to known positions if column names differ
        if reg_col is None:
            reg_col = 0
        if debit_col is None:
            debit_col = 2

        reg_to_debits = {}
        for _, row in df.iterrows():
            reg = str(row[reg_col])
            debit = str(row[debit_col])
            if reg not in reg_to_debits:
                reg_to_debits[reg] = set()
            reg_to_debits[reg].add(debit)

        # Tag style for rows needing attention
        self.tree.tag_configure('highlight', background='#FFFACD')  # Light golden

        for _, row in df.iterrows():
            formatted_row = [str(val) if pd.notna(val) else "" for val in row]
            reg = formatted_row[reg_col]
            # Highlight if this RegNo has more than one distinct DebitNote
            if len(reg_to_debits.get(reg, [])) > 1:
                self.tree.insert("", tk.END, values=formatted_row, tags=('highlight',))
            else:
                self.tree.insert("", tk.END, values=formatted_row)

    # -------------------------------------------------------------
    # Asynchronous Processing
    # -------------------------------------------------------------
    def async_test_connection(self):
        def worker():
            ok, msg = test_connection(
                server=SERVER,
                database=DATABASE,
                username=USERNAME,
                password=PASSWORD,
                auth_type=AUTH_TYPE,
                driver=self.resolved_driver
            )
            def update_ui():
                if ok:
                    self.status_badge.config(text="[ System Ready ]", foreground="#4EC9B0")
                    self.log("System connection ready.", "success")
                else:
                    self.status_badge.config(text="[ Connection Error ]", foreground="#F44747")
                    self.log("System connection error.", "error")
            self.root.after(0, update_ui)

        threading.Thread(target=worker, daemon=True).start()

    def async_process_excel(self):
        start_year = self.start_year_var.get().strip()
        end_year = self.end_year_var.get().strip()

        if not start_year or not end_year:
            messagebox.showwarning("Year Required", "Please select both Start Year and End Year.")
            return

        if int(start_year) > int(end_year):
            messagebox.showwarning("Invalid Range", "Start Year cannot be greater than End Year.")
            return

        # Convert years to date boundaries for SQL query
        start_date = f"{start_year}-01-01"
        end_date = f"{int(end_year) + 1}-01-01"

        excel_path = self.excel_path_var.get().strip()
        if not excel_path:
            messagebox.showwarning("DMVIC Report Required", "Please upload the report from DMVIC.")
            self.browse_excel_file()
            return

        if not os.path.exists(excel_path):
            messagebox.showerror("File Not Found", f"DMVIC report file not found:\n{excel_path}")
            return

        # Default output path
        base, ext = os.path.splitext(excel_path)
        output_path = f"{base}_PURE_DMVIC_RECON{ext}"

        self.set_running_state(True)
        self.log("==========================================", "info")
        self.log(f"Extracting report from: {os.path.basename(excel_path)}", "info")

        def log_cb(msg: str):
            tag = "info"
            if msg.startswith("ERROR"):
                tag = "error"
            elif msg.startswith("SUCCESS"):
                tag = "success"
            elif msg.startswith("WARNING"):
                tag = "warning"
            self.root.after(0, lambda m=msg, t=tag: self.log(m, t))

        def worker():
            ok, df_preview, matched, total, result_msg = process_excel_pure_dmvic_recon(
                excel_path=excel_path,
                output_excel_path=output_path,
                start_date=start_date,
                end_date=end_date,
                server=SERVER,
                database=DATABASE,
                username=USERNAME,
                password=PASSWORD,
                auth_type=AUTH_TYPE,
                driver=self.resolved_driver,
                log_callback=log_cb
            )
            def done_ui():
                self.set_running_state(False)
                if ok:
                    self.processed_df = df_preview
                    self.save_excel_btn.config(state="normal")
                    if df_preview is not None:
                        self.update_treeview(df_preview)
                    self.notebook.select(self.table_tab)
                    messagebox.showinfo(
                        "Extraction Complete",
                        f"Report extracted successfully!\n\nMatched {matched} out of {total} rows.\nHeader set at T4 and DEBITNOTE values populated.\nSaved to:\n{output_path}"
                    )
                else:
                    messagebox.showerror("Extraction Failed", f"Failed to extract report:\n{result_msg}")

            self.root.after(0, done_ui)

        threading.Thread(target=worker, daemon=True).start()

    def save_processed_excel(self):
        excel_path = self.excel_path_var.get().strip()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showwarning("No DMVIC Report", "Please upload a DMVIC report file first.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save Extracted Report",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook (*.xlsx)", "*.xlsx"), ("All Files", "*.*")],
            initialfile="PURE_DMVIC_RECON_Report.xlsx"
        )
        if not file_path:
            return

        start_year = self.start_year_var.get().strip()
        end_year = self.end_year_var.get().strip()
        start_date = f"{start_year}-01-01"
        end_date = f"{int(end_year) + 1}-01-01"
        self.set_running_state(True)

        def log_cb(msg: str):
            self.root.after(0, lambda m=msg: self.log(m, "info"))

        def worker():
            ok, df_preview, matched, total, result_msg = process_excel_pure_dmvic_recon(
                excel_path=excel_path,
                output_excel_path=file_path,
                start_date=start_date,
                end_date=end_date,
                server=SERVER,
                database=DATABASE,
                username=USERNAME,
                password=PASSWORD,
                auth_type=AUTH_TYPE,
                driver=self.resolved_driver,
                log_callback=log_cb
            )
            def done_ui():
                self.set_running_state(False)
                if ok:
                    messagebox.showinfo("Save Successful", f"Saved extracted report to:\n{file_path}")
                else:
                    messagebox.showerror("Save Failed", f"Failed to save report:\n{result_msg}")

            self.root.after(0, done_ui)

        threading.Thread(target=worker, daemon=True).start()

def main():
    root = tk.Tk()
    app = SQLRunnerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
