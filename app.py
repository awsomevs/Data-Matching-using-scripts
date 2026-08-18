import os
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk

from comparator_engine import (
    read_data_file,
    find_target_files,
    compare_datasets,
    export_to_excel,
    copy_matched_files,
    SUPPORTED_EXTENSIONS
)

# Configure CustomTkinter appearance
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class DataComparatorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Config
        self.title("Multi-File Data Comparator (CSV / Excel)")
        self.geometry("1180 x 840")
        self.minsize(980, 700)

        # State Variables
        self.source_file_path = ""
        self.target_files = []
        self.key_checkboxes = {}
        self.comparison_results = None
        self.latest_report_path = ""
        self.copied_matched_folder = ""

        # UI Layout
        self.create_widgets()

    def create_widgets(self):
        # Main Grid Layout (Row 0: Header, Row 1: Content, Row 2: Status)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # -------------------------------------------------------------
        # 1. Header Frame
        # -------------------------------------------------------------
        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=("gray85", "gray17"))
        header_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        header_frame.grid_columnconfigure(0, weight=1)

        header_title = ctk.CTkLabel(
            header_frame,
            text="Multi-File Data Comparator",
            font=ctk.CTkFont(size=22, weight="bold")
        )
        header_title.grid(row=0, column=0, sticky="w", padx=20, pady=(12, 2))

        header_sub = ctk.CTkLabel(
            header_frame,
            text="Compare records from one source CSV/Excel file against single or multiple target files/folders",
            font=ctk.CTkFont(size=12),
            text_color="gray60"
        )
        header_sub.grid(row=1, column=0, sticky="w", padx=20, pady=(0, 12))

        # Appearance mode switch
        self.theme_switch = ctk.CTkSwitch(
            header_frame,
            text="Dark Mode",
            command=self.toggle_theme
        )
        self.theme_switch.grid(row=0, column=1, rowspan=2, padx=20, pady=10, sticky="e")
        self.theme_switch.select()

        # -------------------------------------------------------------
        # 2. Main Content Split View (Left: Input Form, Right: Results)
        # -------------------------------------------------------------
        content_frame = ctk.CTkFrame(self, fg_color="transparent")
        content_frame.grid(row=1, column=0, sticky="nsew", padx=15, pady=15)
        content_frame.grid_columnconfigure(0, weight=4)  # Left panel width
        content_frame.grid_columnconfigure(1, weight=5)  # Right panel width
        content_frame.grid_rowconfigure(0, weight=1)

        # =============================================================
        # LEFT PANEL: Setup Controls
        # =============================================================
        left_panel = ctk.CTkScrollableFrame(content_frame, label_text="Setup Comparison Tasks")
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=0)
        left_panel.grid_columnconfigure(0, weight=1)

        # --- Section 1: Source File ---
        sec1_label = ctk.CTkLabel(left_panel, text="1. Source File Selection", font=ctk.CTkFont(size=14, weight="bold"))
        sec1_label.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))

        source_btn_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        source_btn_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=2)
        source_btn_frame.grid_columnconfigure(0, weight=1)

        self.source_entry = ctk.CTkEntry(source_btn_frame, placeholder_text="Select source CSV or Excel file...")
        self.source_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        btn_browse_source = ctk.CTkButton(source_btn_frame, text="Browse File", width=100, command=self.browse_source_file)
        btn_browse_source.grid(row=0, column=1)

        # --- Section 2: Key Columns ---
        sec2_label = ctk.CTkLabel(left_panel, text="2. Select Key Column(s) to Match On", font=ctk.CTkFont(size=14, weight="bold"))
        sec2_label.grid(row=2, column=0, sticky="w", padx=10, pady=(15, 5))

        self.keys_container = ctk.CTkScrollableFrame(left_panel, height=130, label_text="Available Columns (Select Source File First)")
        self.keys_container.grid(row=3, column=0, sticky="ew", padx=10, pady=2)
        self.keys_container.grid_columnconfigure(0, weight=1)

        # --- Section 3: Target Files / Folder ---
        sec3_label = ctk.CTkLabel(left_panel, text="3. Target File(s) or Directory", font=ctk.CTkFont(size=14, weight="bold"))
        sec3_label.grid(row=4, column=0, sticky="w", padx=10, pady=(15, 5))

        target_btn_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        target_btn_frame.grid(row=5, column=0, sticky="ew", padx=10, pady=2)

        btn_browse_target_dir = ctk.CTkButton(target_btn_frame, text="📁 Browse Folder", width=120, command=self.browse_target_folder)
        btn_browse_target_dir.pack(side="left", padx=(0, 5))

        btn_browse_target_files = ctk.CTkButton(target_btn_frame, text="📄 Select Files", width=120, command=self.browse_target_files)
        btn_browse_target_files.pack(side="left")

        self.chk_recursive_var = ctk.BooleanVar(value=False)
        chk_recursive = ctk.CTkCheckBox(left_panel, text="Search Subfolders Recursively", variable=self.chk_recursive_var, command=self.on_recursive_toggle)
        chk_recursive.grid(row=6, column=0, sticky="w", padx=10, pady=(5, 5))

        self.lbl_target_count = ctk.CTkLabel(left_panel, text="Target Files Selected: 0 files", font=ctk.CTkFont(size=12, weight="bold"), text_color="gray70")
        self.lbl_target_count.grid(row=7, column=0, sticky="w", padx=10, pady=(2, 2))

        self.targets_listbox = ctk.CTkTextbox(left_panel, height=100, font=ctk.CTkFont(size=11))
        self.targets_listbox.grid(row=8, column=0, sticky="ew", padx=10, pady=2)
        self.targets_listbox.insert("1.0", "No target files selected yet...")
        self.targets_listbox.configure(state="disabled")

        # --- Section 4: Output File & File Copy Options ---
        sec4_label = ctk.CTkLabel(left_panel, text="4. Output & Matched File Copy Options", font=ctk.CTkFont(size=14, weight="bold"))
        sec4_label.grid(row=9, column=0, sticky="w", padx=10, pady=(15, 5))

        output_btn_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        output_btn_frame.grid(row=10, column=0, sticky="ew", padx=10, pady=2)
        output_btn_frame.grid_columnconfigure(0, weight=1)

        default_out = os.path.join(os.getcwd(), "comparison_report.xlsx")
        self.output_entry = ctk.CTkEntry(output_btn_frame, placeholder_text="Output Excel report path...")
        self.output_entry.insert(0, default_out)
        self.output_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        btn_browse_output = ctk.CTkButton(output_btn_frame, text="Change", width=80, command=self.browse_output_file)
        btn_browse_output.grid(row=0, column=1)

        # Checkbox to copy matched files
        self.chk_copy_matched_var = ctk.BooleanVar(value=True)
        chk_copy_matched = ctk.CTkCheckBox(
            left_panel,
            text="Copy matched target files to a separate folder",
            variable=self.chk_copy_matched_var,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        chk_copy_matched.grid(row=11, column=0, sticky="w", padx=10, pady=(8, 2))

        # --- Section 5: Run Button ---
        self.btn_run = ctk.CTkButton(
            left_panel,
            text="🚀 START COMPARISON",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=45,
            fg_color="#1F4E78",
            hover_color="#163857",
            command=self.start_comparison_thread
        )
        self.btn_run.grid(row=12, column=0, sticky="ew", padx=10, pady=(20, 15))

        # =============================================================
        # RIGHT PANEL: Results Dashboard & Data View
        # =============================================================
        right_panel = ctk.CTkFrame(content_frame)
        right_panel.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=0)
        right_panel.grid_rowconfigure(1, weight=1)
        right_panel.grid_columnconfigure(0, weight=1)

        # Stats Cards Banner
        self.stats_frame = ctk.CTkFrame(right_panel, fg_color="transparent")
        self.stats_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        self.stats_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.card_total = self.create_stat_card(self.stats_frame, 0, "Total Source", "0")
        self.card_matched = self.create_stat_card(self.stats_frame, 1, "Matched Records", "0 (0%)", color="#2e7d32")
        self.card_unmatched = self.create_stat_card(self.stats_frame, 2, "Unmatched", "0", color="#c62828")
        self.card_targets = self.create_stat_card(self.stats_frame, 3, "Matched Files", "0")

        # Tabview for Results Preview & Actions
        self.tabview = ctk.CTkTabview(right_panel)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        tab_dash = self.tabview.add("Summary & Actions")
        tab_matched = self.tabview.add("Matched Records")
        tab_unmatched = self.tabview.add("Unmatched Records")

        # --- Tab 1: Actions & Details ---
        tab_dash.grid_columnconfigure(0, weight=1)

        action_frame = ctk.CTkFrame(tab_dash, fg_color="transparent")
        action_frame.pack(fill="x", padx=10, pady=10)

        self.btn_open_excel = ctk.CTkButton(
            action_frame,
            text="📊 Open Excel Report",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#2e7d32",
            hover_color="#1b5e20",
            state="disabled",
            command=self.open_excel_report
        )
        self.btn_open_excel.pack(side="left", padx=(0, 8))

        self.btn_open_matched_dir = ctk.CTkButton(
            action_frame,
            text="📂 Open Matched Files Folder",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#1F4E78",
            hover_color="#163857",
            state="disabled",
            command=self.open_matched_folder
        )
        self.btn_open_matched_dir.pack(side="left", padx=(0, 8))

        self.btn_open_folder = ctk.CTkButton(
            action_frame,
            text="📁 Open Report Folder",
            font=ctk.CTkFont(size=13),
            state="disabled",
            command=self.open_output_folder
        )
        self.btn_open_folder.pack(side="left")

        self.log_box = ctk.CTkTextbox(tab_dash, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_box.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        self.log("Ready. Select Source file and Target files to begin.")

        # --- Tab 2: Matched Records Table ---
        self.tree_matched = self.create_treeview(tab_matched)

        # --- Tab 3: Unmatched Records Table ---
        self.tree_unmatched = self.create_treeview(tab_unmatched)

        # -------------------------------------------------------------
        # 3. Bottom Status Bar & Progress Bar
        # -------------------------------------------------------------
        status_bar = ctk.CTkFrame(self, height=35, corner_radius=0)
        status_bar.grid(row=2, column=0, sticky="ew", padx=0, pady=0)
        status_bar.grid_columnconfigure(1, weight=1)

        self.progress_bar = ctk.CTkProgressBar(status_bar, height=8)
        self.progress_bar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(5, 2))
        self.progress_bar.set(0)

        self.status_lbl = ctk.CTkLabel(status_bar, text="Status: Ready", font=ctk.CTkFont(size=11))
        self.status_lbl.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 5))

    def create_stat_card(self, parent, col, title, value, color=None):
        card = ctk.CTkFrame(parent, fg_color=("gray90", "gray22"), corner_radius=8)
        card.grid(row=0, column=col, sticky="ew", padx=4, pady=2)
        card.grid_columnconfigure(0, weight=1)

        lbl_t = ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=11), text_color="gray60")
        lbl_t.grid(row=0, column=0, padx=8, pady=(6, 0), sticky="w")

        lbl_v = ctk.CTkLabel(card, text=value, font=ctk.CTkFont(size=15, weight="bold"))
        if color:
            lbl_v.configure(text_color=color)
        lbl_v.grid(row=1, column=0, padx=8, pady=(0, 6), sticky="w")
        return lbl_v

    def create_treeview(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Style treeview
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Treeview",
            background="#2a2d2e",
            foreground="white",
            rowheight=24,
            fieldbackground="#2a2d2e",
            bordercolor="#2a2d2e"
        )
        style.map('Treeview', background=[('selected', '#1F4E78')])
        style.configure("Treeview.Heading", background="#1F4E78", foreground="white", font=('Calibri', 10, 'bold'))

        tree = ttk.Treeview(frame, show="headings")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.pack(side="left", fill="both", expand=True)

        return tree

    def log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def toggle_theme(self):
        if self.theme_switch.get() == 1:
            ctk.set_appearance_mode("Dark")
        else:
            ctk.set_appearance_mode("Light")

    # -------------------------------------------------------------
    # Event Handlers & File Browsing
    # -------------------------------------------------------------
    def browse_source_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Source Data File",
            filetypes=[("Tabular Files", "*.csv *.xlsx *.xls *.tsv"), ("All Files", "*.*")]
        )
        if file_path:
            self.source_file_path = os.path.abspath(file_path)
            self.source_entry.delete(0, "end")
            self.source_entry.insert(0, self.source_file_path)
            self.load_source_columns()

    def load_source_columns(self):
        if not self.source_file_path or not os.path.exists(self.source_file_path):
            return

        try:
            df = read_data_file(self.source_file_path)
            columns = list(df.columns)

            # Clear existing checkboxes
            for widget in self.keys_container.winfo_children():
                widget.destroy()

            self.key_checkboxes = {}

            if not columns:
                lbl = ctk.CTkLabel(self.keys_container, text="No columns found in source file.")
                lbl.pack(padx=5, pady=5)
                return

            lbl_info = ctk.CTkLabel(self.keys_container, text="Check matching column(s):", font=ctk.CTkFont(size=11, weight="bold"))
            lbl_info.pack(anchor="w", padx=5, pady=(2, 5))

            for col in columns:
                var = ctk.BooleanVar(value=False)
                # Auto check if column name looks like account/phone/id/complaint
                col_lower = col.lower()
                if any(kw in col_lower for kw in ['account', 'phone', 'mobile', 'complaint', 'acknowledgement', 'id', 'number']):
                    var.set(True)

                chk = ctk.CTkCheckBox(self.keys_container, text=col, variable=var)
                chk.pack(anchor="w", padx=10, pady=2)
                self.key_checkboxes[col] = var

            self.log(f"Loaded source file '{os.path.basename(self.source_file_path)}' with {len(columns)} columns and {len(df)} rows.")

        except Exception as e:
            messagebox.showerror("Source File Error", f"Failed to load source file:\n{str(e)}")
            self.log(f"[ERROR] Source file load error: {str(e)}")

    def browse_target_folder(self):
        folder_path = filedialog.askdirectory(title="Select Target Folder")
        if folder_path:
            target_files = find_target_files(folder_path, recursive=self.chk_recursive_var.get())
            self.set_target_files(target_files)

    def browse_target_files(self):
        files = filedialog.askopenfilenames(
            title="Select Target Files",
            filetypes=[("Tabular Files", "*.csv *.xlsx *.xls *.tsv"), ("All Files", "*.*")]
        )
        if files:
            self.set_target_files(list(files))

    def on_recursive_toggle(self):
        if self.target_files and os.path.isdir(os.path.dirname(self.target_files[0])):
            parent_dir = os.path.dirname(self.target_files[0])
            target_files = find_target_files(parent_dir, recursive=self.chk_recursive_var.get())
            self.set_target_files(target_files)

    def set_target_files(self, files):
        source_abs = os.path.abspath(self.source_file_path) if self.source_file_path else ""
        self.target_files = [f for f in files if os.path.abspath(f) != source_abs]

        self.lbl_target_count.configure(text=f"Target Files Selected: {len(self.target_files)} files")

        self.targets_listbox.configure(state="normal")
        self.targets_listbox.delete("1.0", "end")
        if self.target_files:
            for f in self.target_files:
                self.targets_listbox.insert("end", f"• {os.path.basename(f)}  ({f})\n")
        else:
            self.targets_listbox.insert("1.0", "No target files found in selected path.")
        self.targets_listbox.configure(state="disabled")

        self.log(f"Selected {len(self.target_files)} target file(s) for comparison.")

    def browse_output_file(self):
        out_path = filedialog.asksaveasfilename(
            title="Save Output Excel Report As",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if out_path:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, out_path)

    # -------------------------------------------------------------
    # Comparison Execution Thread
    # -------------------------------------------------------------
    def start_comparison_thread(self):
        source_path = self.source_entry.get().strip()
        if not source_path or not os.path.exists(source_path):
            messagebox.showwarning("Missing Source", "Please select a valid source CSV or Excel file.")
            return

        if not self.target_files:
            messagebox.showwarning("Missing Targets", "Please select target folder or target files to compare against.")
            return

        selected_keys = [col for col, var in self.key_checkboxes.items() if var.get()]
        if not selected_keys:
            messagebox.showwarning("No Key Columns", "Please check at least one Key Column to match records on.")
            return

        output_path = self.output_entry.get().strip()
        if not output_path:
            messagebox.showwarning("Missing Output Path", "Please enter output file path.")
            return

        should_copy_matched = self.chk_copy_matched_var.get()

        # Disable UI elements during run
        self.btn_run.configure(state="disabled", text="⏳ Comparing Data...")
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()
        self.status_lbl.configure(text="Status: Comparing datasets across target files...")

        self.log("\n==================================================")
        self.log(f"Starting comparison run on {len(self.target_files)} target file(s)...")
        self.log(f"Key Column(s): {selected_keys}")

        # Launch background thread
        threading.Thread(
            target=self.run_comparison,
            args=(source_path, self.target_files, selected_keys, output_path, should_copy_matched),
            daemon=True
        ).start()

    def run_comparison(self, source_path, target_files, selected_keys, output_path, should_copy_matched):
        try:
            results = compare_datasets(source_path, target_files, selected_keys)
            saved_excel_path = export_to_excel(results, output_path)

            copied_matched_dir = ""
            copied_count = 0

            # Copy matched target files if option enabled
            if should_copy_matched and results.get('matched_target_paths'):
                out_dir = os.path.dirname(os.path.abspath(output_path))
                base_name = os.path.splitext(os.path.basename(output_path))[0]
                copied_matched_dir = os.path.join(out_dir, f"{base_name}_Matched_Files")

                copied_files = copy_matched_files(results['matched_target_paths'], copied_matched_dir)
                copied_count = len(copied_files)

            self.comparison_results = results
            self.latest_report_path = saved_excel_path
            self.copied_matched_folder = copied_matched_dir

            # Update UI on main thread
            self.after(0, self.on_comparison_success, results, saved_excel_path, copied_matched_dir, copied_count)

        except Exception as e:
            self.after(0, self.on_comparison_error, str(e))

    def on_comparison_success(self, results, saved_excel_path, copied_matched_dir, copied_count):
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(1.0)
        self.btn_run.configure(state="normal", text="🚀 START COMPARISON")
        self.status_lbl.configure(text="Status: Comparison Complete! Matched files copied.")

        stats = results['summary_stats']

        # Update Stat Cards
        self.card_total.configure(text=str(stats['Total Source Records']))
        self.card_matched.configure(text=f"{stats['Found in Targets (Matched)']} ({stats['Match Percentage']})")
        self.card_unmatched.configure(text=str(stats['Missing in Targets (Unmatched)']))
        self.card_targets.configure(text=f"{stats['Matched Target Files Count']} / {stats['Total Target Files Checked']}")

        # Enable action buttons
        self.btn_open_excel.configure(state="normal")
        self.btn_open_folder.configure(state="normal")

        if copied_matched_dir and os.path.exists(copied_matched_dir):
            self.btn_open_matched_dir.configure(state="normal")
        else:
            self.btn_open_matched_dir.configure(state="disabled")

        # Update Log
        self.log("[SUMMARY] COMPARISON COMPLETED SUCCESSFULLY!")
        self.log(f"• Total Source Records:      {stats['Total Source Records']}")
        self.log(f"• Matched Source Records:    {stats['Found in Targets (Matched)']} ({stats['Match Percentage']})")
        self.log(f"• Unmatched Source Records:  {stats['Missing in Targets (Unmatched)']}")
        self.log(f"• Matched Target Files:      {stats['Matched Target Files Count']} file(s)")
        self.log(f"• Excel Report Saved To:     {saved_excel_path}")

        if copied_matched_dir and copied_count > 0:
            self.log(f"• Matched Target Files Copied ({copied_count} files) To:\n  {copied_matched_dir}")

        # Populate Treeviews
        self.populate_treeview(self.tree_matched, results['matched_df'])
        self.populate_treeview(self.tree_unmatched, results['unmatched_df'])

        msg = (
            f"Comparison completed successfully!\n\n"
            f"• Matched Records: {stats['Found in Targets (Matched)']} / {stats['Total Source Records']} ({stats['Match Percentage']})\n"
            f"• Matched Target Files: {stats['Matched Target Files Count']} file(s)\n\n"
            f"Excel Report Saved To:\n{saved_excel_path}"
        )
        if copied_count > 0:
            msg += f"\n\n📂 Copies of {copied_count} matched target file(s) saved to:\n{copied_matched_dir}"

        messagebox.showinfo("Comparison Complete", msg)

    def on_comparison_error(self, err_msg):
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(0)
        self.btn_run.configure(state="normal", text="🚀 START COMPARISON")
        self.status_lbl.configure(text="Status: Error during comparison!")

        self.log(f"\n[ERROR] Comparison failed: {err_msg}")
        messagebox.showerror("Comparison Error", f"An error occurred during comparison:\n\n{err_msg}")

    def populate_treeview(self, tree, df):
        tree.delete(*tree.get_children())
        tree["columns"] = []

        if df.empty:
            return

        cols = list(df.columns)
        tree["columns"] = cols

        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=120, minwidth=80, anchor="w")

        for idx, row in df.head(500).iterrows():
            vals = [row[c] for c in cols]
            tree.insert("", "end", values=vals)

    def open_excel_report(self):
        if self.latest_report_path and os.path.exists(self.latest_report_path):
            os.startfile(self.latest_report_path)

    def open_matched_folder(self):
        if self.copied_matched_folder and os.path.exists(self.copied_matched_folder):
            os.startfile(self.copied_matched_folder)

    def open_output_folder(self):
        if self.latest_report_path:
            folder = os.path.dirname(self.latest_report_path)
            if os.path.exists(folder):
                os.startfile(folder)


if __name__ == "__main__":
    app = DataComparatorApp()
    app.mainloop()
