# ui_extract_view.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
import threading
import os
from typing import List, Dict, Any, Callable

from ...classes.pq_manager import PQManager
from ...classes.pq_manager.models import PowerQueryScript, PowerQueryMetadata
from ..theme import SoP


class ExtractView(ctk.CTkFrame):
    """
    Manages the 'Extract' view for pulling queries from Excel.
    """

    def __init__(self, parent, manager: PQManager, refresh_callback: Callable):
        super().__init__(parent, fg_color="transparent")
        self.manager = manager
        self.refresh_callback = refresh_callback

        self.excel_file_to_extract = ""
        self.extraction_vars = []  # For the confirmation dialog

        self._build_widgets()

    def _build_widgets(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)  # Log box

        title = ctk.CTkLabel(
            self,
            text="Extract Queries from Excel",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=SoP["ACCENT_HOVER"]
        )
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))

        # --- Options Panel ---
        top_panel = ctk.CTkFrame(
            self, fg_color=SoP["FRAME"], corner_radius=8)
        top_panel.grid(row=1, column=0, sticky="new", pady=10)
        top_panel.grid_columnconfigure(0, weight=1)
        top_panel.grid_columnconfigure(1, weight=1)

        # --- Option 1: Active Workbook ---
        self.extract_active_btn = ctk.CTkButton(
            top_panel,
            text="📥 Extract from Active Workbook",
            height=50,
            command=self._threaded_get_queries_active,
            fg_color=SoP["ACCENT"],
            hover_color=SoP["ACCENT_HOVER"],
            text_color="#000000",
            font=ctk.CTkFont(weight="bold")
        )
        self.extract_active_btn.grid(
            row=0, column=0, sticky="ew", padx=15, pady=15)

        # --- Option 2: Select File ---
        self.extract_file_btn = ctk.CTkButton(
            top_panel,
            text="📁 Select File...",
            height=50,
            command=self._select_excel_file,
            fg_color="transparent",
            border_width=1,
            border_color=SoP["ACCENT"],
            hover_color=SoP["TREE_FIELD"],
            text_color=SoP["ACCENT"]
        )
        self.extract_file_btn.grid(
            row=1, column=0, sticky="ew", padx=15, pady=(0, 15))

        self.extract_file_label = ctk.CTkLabel(
            top_panel,
            text="No file selected.",
            text_color=SoP["TEXT_DIM"]
        )
        self.extract_file_label.grid(row=1, column=1, sticky="w", padx=20)

        self.extract_from_file_start_btn = ctk.CTkButton(
            top_panel,
            text="Get Queries from File",
            height=40,
            command=self._threaded_get_queries_from_file,
            fg_color=SoP["FRAME"],
            hover_color=SoP["TREE_FIELD"],
            text_color=SoP["TEXT_DIM"],
            state="disabled"
        )
        self.extract_from_file_start_btn.grid(
            row=2, column=0, columnspan=2, sticky="ew", padx=15, pady=(0, 15))

        # --- Log/Output Box ---
        ctk.CTkLabel(self, text="Extraction Log", text_color=SoP["TEXT_DIM"]).grid(
            row=2, column=0, sticky="sw", pady=(10, 5))

        self.extract_log = ctk.CTkTextbox(
            self,
            border_color=SoP["FRAME"],
            border_width=2,
            fg_color=SoP["FRAME"],
            text_color=SoP["TEXT_DIM"]
        )
        self.extract_log.insert("1.0", "Extraction log will appear here...")
        self.extract_log.configure(state="disabled")
        self.extract_log.grid(row=3, column=0, columnspan=2,
                              sticky="nsew", pady=(0, 10))

        # Configure log tags
        self.extract_log.tag_config("accent", foreground=SoP["ACCENT"])
        self.extract_log.tag_config(
            "accent_bold", foreground=SoP["ACCENT_HOVER"])
        self.extract_log.tag_config("error", foreground="#FF5555")

    # --- Extract from Excel Handlers ---

    def _select_excel_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel Files", "*.xlsx;*.xlsm;*.xlsb"), ("All Files", "*.*")))
        if path:
            self.excel_file_to_extract = path
            self.extract_file_label.configure(text=os.path.basename(path))
            self.extract_from_file_start_btn.configure(
                state="normal", text_color=SoP["ACCENT"])

    def _clear_log(self):
        self.extract_log.configure(state="normal")
        self.extract_log.delete("1.0", "end")
        self.extract_log.configure(state="disabled")

    def _append_log(self, message, tag=None):
        def _task():
            self.extract_log.configure(state="normal")
            if tag:
                self.extract_log.insert("end", message, (tag,))
            else:
                self.extract_log.insert("end", message)
            self.extract_log.see("end")
            self.extract_log.configure(state="disabled")
        self.master.after(0, _task)

    def _threaded_get_queries_active(self):
        """Step 1: Get queries from active workbook."""
        self._clear_log()
        self._append_log(
            "Attempting to read from active Excel workbook...\n\n", "accent")
        self.extract_active_btn.configure(state="disabled", text="Reading...")

        threading.Thread(
            target=self._get_queries,
            args=(None, "Active Workbook"),  # None means active
            daemon=True
        ).start()

    def _threaded_get_queries_from_file(self):
        """Step 1: Get queries from selected file."""
        if not self.excel_file_to_extract:
            messagebox.showwarning(
                "No File", "Please select an Excel file first.")
            return

        self._clear_log()
        self._append_log(
            f"Starting extraction from {self.excel_file_to_extract}...\n\n", "accent")
        self.extract_from_file_start_btn.configure(
            state="disabled", text="Reading...")

        threading.Thread(
            target=self._get_queries,
            args=(self.excel_file_to_extract, os.path.basename(
                self.excel_file_to_extract)),
            daemon=True
        ).start()

    def _get_queries(self, file_path: str | None, source_name: str):
        """
        Generic function to *get* queries using the new Excel service.
        `file_path` is None for active, or a string path.
        `source_name` is just for display.
        """
        try:
            # NEW API CALL:
            query_list = self.manager.excel.get_queries_from_workbook(
                file_path)

            if not query_list:
                self._append_log(
                    "No Power Queries found in the workbook.", "error")
                messagebox.showinfo(
                    "No Queries Found", f"No Power Queries were found in {source_name}.")
                return

            self._append_log(
                f"Found {len(query_list)} queries. Opening confirmation dialog...", "accent")
            # Open the confirmation dialog on the main thread
            self.master.after(0, lambda: self._open_extraction_confirmation_dialog(
                query_list, source_name))

        except Exception as e:
            self._append_log(f"\n--- ERROR ---\n{e}\n", "error")
            self.master.after(0, lambda: messagebox.showerror(
                "Extraction Error", f"An error occurred while reading the file:\n{e}"))
        finally:
            # Re-enable buttons
            def _reset():
                self.extract_active_btn.configure(
                    state="normal", text="📥 Extract from Active Workbook")
                self.extract_from_file_start_btn.configure(
                    state="normal", text="Get Queries from File")
            self.master.after(0, _reset)

    # --- Extraction Confirmation Dialog ---

    def _open_extraction_confirmation_dialog(self, query_list: List[Dict[str, Any]], source_name: str):
        """Shows a new window to let the user pick which queries to import."""

        dialog = ctk.CTkToplevel(self)
        dialog.title("Confirm Queries to Import")
        dialog.geometry("700x600")
        dialog.transient()
        dialog.grab_set()
        dialog.configure(fg_color=SoP["BG"])
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(2, weight=1)

        self.extraction_vars = []  # Clear previous vars

        title_label = ctk.CTkLabel(
            dialog,
            text=f"Found {len(query_list)} queries in {source_name}",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=SoP["ACCENT_HOVER"]
        )
        title_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        # ... (Rest of dialog layout is identical to your original code) ...
        # --- Search bar for the dialog ---
        search_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        search_frame.grid(row=1, column=0, sticky="ew", padx=15)
        search_frame.grid_columnconfigure(0, weight=1)

        dialog_search_var = tk.StringVar()
        dialog_search_entry = ctk.CTkEntry(
            search_frame,
            textvariable=dialog_search_var,
            placeholder_text="Filter queries...",
            fg_color=SoP["TREE_FIELD"],
            border_color=SoP["FRAME"],
            text_color=SoP["TEXT"],
            height=35
        )
        dialog_search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        # --- Checkbox list ---
        scroll_frame = ctk.CTkScrollableFrame(
            dialog, fg_color=SoP["FRAME"], corner_radius=8
        )
        scroll_frame.grid(row=2, column=0, sticky="nsew", padx=15, pady=10)
        scroll_frame.grid_columnconfigure(0, weight=1)

        # Function to filter checkboxes
        def filter_queries(*args):
            query = dialog_search_var.get().lower()
            for var_tuple in self.extraction_vars:
                cb, _, label, query_dict = var_tuple
                name = query_dict.get("name", "").lower()
                desc = query_dict.get("description", "").lower()

                if query in name or query in desc:
                    cb.grid()
                    label.grid()
                else:
                    cb.grid_remove()
                    label.grid_remove()

        dialog_search_var.trace_add("write", filter_queries)

        for i, query in enumerate(query_list):
            var = tk.BooleanVar(value=True)
            name = str(query.get("name"))
            desc = (query.get("description")
                    or "No description.")[:100] + "..."

            cb = ctk.CTkCheckBox(
                scroll_frame,
                text=name,
                variable=var,
                fg_color=SoP["ACCENT"],
                hover_color=SoP["ACCENT_HOVER"],
                text_color=SoP["TEXT"]
            )
            cb.grid(row=i*2, column=0, sticky="w", padx=10, pady=(10, 0))

            label = ctk.CTkLabel(
                scroll_frame,
                text=f"   {desc}",
                text_color=SoP["TEXT_DIM"]
            )
            label.grid(row=i*2 + 1, column=0, sticky="w", padx=20, pady=(0, 5))

            # Store widgets for filtering
            self.extraction_vars.append((cb, var, label, query))

        # --- Action Buttons ---
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=10)
        btn_frame.grid_columnconfigure(3, weight=1)

        def select_all(val):
            for var_tuple in self.extraction_vars:
                cb, var, label, _ = var_tuple
                if cb.winfo_viewable():  # Only affect visible items
                    var.set(val)

        ctk.CTkButton(
            btn_frame, text="Select All Visible", fg_color=SoP["FRAME"], text_color=SoP["TEXT_DIM"],
            command=lambda: select_all(True)
        ).grid(row=0, column=0, padx=5)

        ctk.CTkButton(
            btn_frame, text="Deselect All Visible", fg_color=SoP["FRAME"], text_color=SoP["TEXT_DIM"],
            command=lambda: select_all(False)
        ).grid(row=0, column=1, padx=5)

        import_btn = ctk.CTkButton(
            btn_frame, text=f"Import {len(query_list)} Queries",
            command=lambda: self._threaded_confirm_extraction(dialog),
            fg_color=SoP["ACCENT"], hover_color=SoP["ACCENT_HOVER"],
            text_color="#000000", font=ctk.CTkFont(weight="bold"), height=40
        )
        import_btn.grid(row=0, column=3, sticky="e")

        # Update button text on selection change
        def update_btn_text(*args):
            count = sum(
                1 for _, var, _, _ in self.extraction_vars if var.get())
            import_btn.configure(text=f"Import {count} Queries")

        for _, var, _, _ in self.extraction_vars:
            var.trace_add("write", update_btn_text)

    def _threaded_confirm_extraction(self, dialog: ctk.CTkToplevel):
        """Step 2: Gathers selected queries and passes them to the writer thread."""
        selected_queries = [query for _, var, _,
                            query in self.extraction_vars if var.get()]

        if not selected_queries:
            messagebox.showwarning(
                "No Queries Selected", "Please select at least one query to import.")
            return

        dialog.destroy()

        self._clear_log()
        self._append_log(
            f"Importing {len(selected_queries)} selected queries...\n\n", "accent")

        threading.Thread(
            target=self._confirm_extraction,
            args=(selected_queries,),
            daemon=True
        ).start()

    def _confirm_extraction(self, selected_queries: List[Dict[str, Any]]):
        """Step 3: (Thread Target) Creates the files and refreshes the UI."""
        created_files = []
        try:
            # NEW API LOGIC: Loop, create script objects, and save.
            for q in selected_queries:
                try:
                    name = q["name"]
                    category = "Extracted"

                    # 1. Determine safe path
                    safe_name = "".join(c for c in name if c.isalnum()
                                        or c in (' ', '_', '-')).rstrip()
                    safe_category = "Extracted"

                    target_dir = os.path.join(
                        self.manager.store.root, safe_category)
                    out_path = os.path.join(target_dir, f"{safe_name}.pq")

                    # 2. Create Pydantic Models
                    meta = PowerQueryMetadata(
                        name=name,
                        category=category,
                        description=q["description"],
                        tags=["extracted"],
                        dependencies=[],  # We don't know deps on extraction
                        version="1.0",
                        path=out_path
                    )
                    script = PowerQueryScript(meta=meta, body=q["formula"])

                    # 3. Save using the store
                    self.manager.store.save_script(script, overwrite=True)
                    created_files.append(out_path)

                except Exception as e:
                    self._append_log(
                        f"Failed to save query {q['name']}: {e}\n", "error")

            # 4. Rebuild the index *once* after all files are saved
            if created_files:
                self._append_log("Rebuilding file index...\n", "accent")
                self.manager.build_index()

            self._append_log(
                f"\nImport complete.\nCreated {len(created_files)} files.\n", "accent_bold")

            # 5. Use callback to refresh the main UI
            self.master.after(0, self.refresh_callback)

        except Exception as e:
            self._append_log(f"\n--- FATAL ERROR ---\n{e}\n", "error")
            self.master.after(0, lambda: messagebox.showerror(
                "Import Error", f"An error occurred:\n{e}"))
