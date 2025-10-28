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
from ..components.codeview import CTkCodeView


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
        self.open_workbook_names = []

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

        # --- Option 1: Open Workbook (NEW) ---
        wb_frame = ctk.CTkFrame(top_panel, fg_color="transparent")
        wb_frame.grid(row=0, column=0, columnspan=2,
                      sticky="ew", padx=15, pady=15)
        wb_frame.grid_columnconfigure(0, weight=1)

        self.workbook_menu = ctk.CTkOptionMenu(
            wb_frame,
            values=["Click 'Refresh' to list..."],
            fg_color=SoP["EDITOR"],
            button_color=SoP["ACCENT_DARK"],
            button_hover_color=SoP["ACCENT"],
            text_color=SoP["TEXT_DIM"],
            height=40
        )
        self.workbook_menu.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.refresh_wbs_btn = ctk.CTkButton(
            wb_frame,
            text="🔄",
            width=40,
            height=40,
            font=ctk.CTkFont(size=20),
            command=self._refresh_workbook_list,
            fg_color=SoP["TREE_FIELD"],
            hover_color=SoP["ACCENT_HOVER"]
        )
        self.refresh_wbs_btn.grid(row=0, column=1, padx=5)

        self.extract_from_wb_btn = ctk.CTkButton(
            wb_frame,
            text="📥 Extract from Selected",
            height=40,
            command=self._threaded_get_queries_from_workbook,
            fg_color=SoP["ACCENT"],
            hover_color=SoP["ACCENT_HOVER"],
            text_color="#000000",
            font=ctk.CTkFont(weight="bold")
        )
        self.extract_from_wb_btn.grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))

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
    def _refresh_workbook_list(self):
        """Calls the manager to get open workbooks and updates the dropdown."""
        try:
            # <-- Gets list[str]
            self.open_workbook_names = self.manager.excel.list_open_workbooks()

            if not self.open_workbook_names:
                self.workbook_menu.configure(
                    values=["No workbooks open."], state="disabled")
                self.workbook_menu.set("No workbooks open.")
                self.extract_from_wb_btn.configure(state="disabled")
            else:
                self.workbook_menu.configure(
                    values=self.open_workbook_names, state="normal")
                self.workbook_menu.set(self.open_workbook_names[0])
                self.extract_from_wb_btn.configure(state="normal")
        except Exception as e:
            messagebox.showerror(
                "Error", f"Failed to list open workbooks:\n{e}")

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

    def _threaded_get_queries_from_workbook(self):
        """Step 1: Get queries from the selected open workbook."""
        selected_name = self.workbook_menu.get()

        if not selected_name or selected_name == "No workbooks open.":
            messagebox.showwarning(
                "No Workbook", "Please select a valid workbook.")
            return

        self._clear_log()
        self._append_log(
            f"Attempting to read from open workbook: {selected_name}...\n\n", "accent")
        self.extract_from_wb_btn.configure(state="disabled", text="Reading...")

        # We pass ONLY the string name to the new thread
        threading.Thread(
            target=self._get_queries_open_wb_target,  # <-- New thread target
            args=(selected_name,),
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

        file_path = self.excel_file_to_extract
        source_name = os.path.basename(file_path)

        # We pass ONLY strings to the new thread
        threading.Thread(
            target=self._get_queries_file_target,  # <-- New thread target
            args=(file_path, source_name),
            daemon=True
        ).start()

    def _get_queries_open_wb_target(self, workbook_name: str):
        """
        THREAD TARGET: Gets queries from an open workbook.
        Runs on a worker thread.
        """
        try:
            # This acquires AND uses the COM object on the same thread
            query_list = self.manager.excel.get_queries_from_open_workbook(
                workbook_name)

            if not query_list:
                self._append_log(
                    "No Power Queries found in the workbook.", "error")
                messagebox.showinfo(
                    "No Queries Found", f"No Power Queries were found in {workbook_name}.")
                return

            self._append_log(
                f"Found {len(query_list)} queries. Opening confirmation dialog...", "accent")
            self.master.after(0, lambda: self._open_extraction_confirmation_dialog(
                query_list, workbook_name))

        except Exception as e:
            self._append_log(f"\n--- ERROR ---\n{e}\n", "error")
            self.master.after(0, lambda: messagebox.showerror(
                "Extraction Error", f"An error occurred:\n{e}"))
        finally:
            self.master.after(0, lambda: self.extract_from_wb_btn.configure(
                state="normal", text="📥 Extract from Selected"))

    def _get_queries_file_target(self, file_path: str, source_name: str):
        """
        THREAD TARGET: Gets queries from a closed file.
        Runs on a worker thread.
        """
        try:
            # This creates a new Excel instance on this thread
            query_list = self.manager.excel.get_queries_from_workbook(
                file_path=file_path)

            if not query_list:
                self._append_log(
                    "No Power Queries found in the workbook.", "error")
                messagebox.showinfo(
                    "No Queries Found", f"No Power Queries were found in {source_name}.")
                return

            self._append_log(
                f"Found {len(query_list)} queries. Opening confirmation dialog...", "accent")
            self.master.after(0, lambda: self._open_extraction_confirmation_dialog(
                query_list, source_name))

        except Exception as e:
            self._append_log(f"\n--- ERROR ---\n{e}\n", "error")
            self.master.after(0, lambda: messagebox.showerror(
                "Extraction Error", f"An error occurred:\n{e}"))
        finally:
            self.master.after(0, lambda: self.extract_from_file_start_btn.configure(
                state="normal", text="Get Queries from File"))

    def _get_queries(self, file_path: str | None, source_name: str, wb_api: Any | None):
        """
        Generic function to *get* queries using the new Excel service.
        `file_path` is None for active, or a string path.
        `source_name` is just for display.
        """
        try:
            # NEW API CALL:
            query_list = self.manager.excel.get_queries_from_workbook(
                file_path=file_path,
                wb_api=wb_api)

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
                self.extract_from_wb_btn.configure(
                    state="normal", text="📥 Extract from Selected")
                self.extract_from_file_start_btn.configure(
                    state="normal", text="Get Queries from File")
            self.master.after(0, _reset)

    # --- Extraction Confirmation Dialog ---
    def _clear_all_tabs(self, tabs: Dict[str, ctk.CTkTextbox]):
        """Clears all textboxes in the tab view."""
        for name, box in tabs.items():
            box.configure(state="normal")
            box.delete("1.0", "end")
            if name == "preview":
                box.insert(
                    "1.0", "Select a query name to preview its M-code.", ("dim",))
            elif name == "params":
                box.insert(
                    "1.0", "Select a query name to view its parameters.", ("dim",))
            elif name == "sources":
                box.insert(
                    "1.0", "Select a query name to view its data sources.", ("dim",))
            box.configure(state="disabled")

    def _open_extraction_confirmation_dialog(self, query_list: List[Dict[str, Any]], source_name: str):
        """
        Shows a new, ADVANCED dialog to let the user preview
        and select which queries to import.
        """
        self.extraction_vars = []  # Clear previous vars

        dialog = ctk.CTkToplevel(self)
        dialog.title(f"Extract Queries from {source_name}")
        dialog.geometry("1100x750")  # Make it bigger
        dialog.transient()
        dialog.grab_set()
        dialog.configure(fg_color=SoP["BG"])
        dialog.grid_columnconfigure(1, weight=3)  # Right panel (preview)
        dialog.grid_columnconfigure(0, weight=1)  # Left panel (list)
        dialog.grid_rowconfigure(1, weight=1)

        title_label = ctk.CTkLabel(
            dialog,
            text=f"Found {len(query_list)} queries. Select queries to import.",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=SoP["ACCENT_HOVER"]
        )
        title_label.grid(row=0, column=0, columnspan=2,
                         padx=20, pady=(20, 10), sticky="w")

        # --- Left Panel: Query List ---
        left_panel = ctk.CTkFrame(dialog, fg_color=SoP["FRAME"])
        left_panel.grid(row=1, column=0, sticky="nsew", padx=(15, 10), pady=10)
        left_panel.grid_columnconfigure(0, weight=1)
        left_panel.grid_rowconfigure(1, weight=1)

        # Search/Filter Bar
        dialog_search_var = tk.StringVar()
        dialog_search_entry = ctk.CTkEntry(
            left_panel,
            textvariable=dialog_search_var,
            placeholder_text="Filter queries...",
            fg_color=SoP["EDITOR"], border_color=SoP["TREE_FIELD"],
            text_color=SoP["TEXT"], height=35
        )
        dialog_search_entry.grid(
            row=0, column=0, sticky="ew", padx=10, pady=10)

        # Scrollable list for checkboxes
        scroll_frame = ctk.CTkScrollableFrame(
            left_panel, fg_color="transparent", corner_radius=8
        )
        scroll_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=(0, 10))
        scroll_frame.grid_columnconfigure(1, weight=1)

        # --- Right Panel: TabView for Previews ---
        right_panel = ctk.CTkTabview(
            dialog,
            fg_color=SoP["FRAME"],
            segmented_button_fg_color=SoP["EDITOR"],
            segmented_button_selected_color=SoP["ACCENT_DARK"],
            segmented_button_selected_hover_color=SoP["ACCENT_DARK"],
            segmented_button_unselected_color=SoP["EDITOR"],
            segmented_button_unselected_hover_color=SoP["TREE_FIELD"],
            text_color=SoP["TEXT_DIM"],
            border_width=1,
            border_color=SoP["FRAME"]
        )
        right_panel.grid(row=1, column=1, sticky="nsew", padx=(0, 15), pady=10)

        tab_preview = CTkCodeView(
            right_panel.add("Preview"),
            manager=self.manager
        )
        tab_params = ctk.CTkTextbox(right_panel.add(
            "Parameters"), fg_color="transparent",  font=("Consolas", 12))
        tab_sources = ctk.CTkTextbox(right_panel.add(
            "Data Sources"), fg_color="transparent", font=("Consolas", 12))

        tab_widgets = {"preview": tab_preview,
                       "params": tab_params, "sources": tab_sources}
        for name, box in tab_widgets.items():
            box.pack(fill="both", expand=True)
            if name == "preview":
                pass  # CTkCodeView handles its tags
            else:
                # Configure the standard CTkTextboxes
                box.tag_config("dim", foreground=SoP["TEXT_DIM"])
                if name == "params":
                    box.tag_config("name", foreground=SoP["ACCENT_HOVER"])
                    box.tag_config("type", foreground="#ce9178")  # Orange
                elif name == "sources":
                    # Mauve (like code)
                    box.tag_config("type", foreground="#c586c0")
                    box.tag_config("source", foreground="#ce9178")  # Orange

        tab_widgets = {"preview": tab_preview,
                       "params": tab_params, "sources": tab_sources}
        for box in tab_widgets.values():
            box.pack(fill="both", expand=True)

        def clear_tabs():
            for name, box in tab_widgets.items():
                box.configure(state="normal")
                box.delete("1.0", "end")
                # Insert placeholder text
                if name == "preview":
                    box.insert("1.0", "Select query name...", ("dim",))
                elif name == "params":
                    box.insert("1.0", "Select query name...", ("dim",))
                elif name == "sources":
                    box.insert("1.0", "Select query name...", ("dim",))
                box.configure(state="disabled")

        clear_tabs()

        # --- Helper function to update the preview panel ---
        def update_preview_panel(query_dict: Dict[str, Any]):
            body = query_dict.get("formula", "")

            # 1. Update Preview Tab
            tab_preview.set_code(body)

            # 2. Update Parameters Tab
            params = self.manager.get_parameters_from_code(body)
            tab_params.configure(state="normal")
            tab_params.delete("1.0", "end")
            if not params:
                tab_params.insert(
                    "1.0", "This query is not a function.", ("dim",))
            else:
                for p in params:
                    tab_params.insert("end", f"Name:     ", ("dim",))
                    tab_params.insert("end", f"{p['name']}\n", ("name",))
                    tab_params.insert("end", f"Type:     ", ("dim",))
                    tab_params.insert("end", f"{p['type']}\n", ("type",))
                    tab_params.insert("end", f"Optional: ", ("dim",))
                    tab_params.insert(
                        "end", f"{'Yes' if p['optional'] else 'No'}\n\n", ("name",))
            tab_params.configure(state="disabled")

            # 3. Update Data Sources Tab
            sources = self.manager.get_datasources_from_code(body)
            tab_sources.configure(state="normal")
            tab_sources.delete("1.0", "end")
            if not sources:
                tab_sources.insert(
                    "1.0", "No external data sources found.", ("dim",))
            else:
                for src in sources:
                    tab_sources.insert("end", f"Type:   ", ("dim",))
                    tab_sources.insert("end", f"{src['type']}\n", ("type",))
                    tab_sources.insert("end", f"Source: ", ("dim",))
                    tab_sources.insert(
                        "end", f"{src['full_argument']}", ("source",))
                    tab_sources.insert(
                        "end", f" (Input Parameter)\n\n" if src["source_type"] == "Variable" else "\n\n")
            tab_sources.configure(state="disabled")

        # --- Populate the Checkbox List ---

        # This will hold (checkbox, button, query_dict)
        query_widgets = []

        for i, query in enumerate(query_list):
            var = tk.BooleanVar(value=True)
            name = str(query.get("name"))

            # Create a frame for each row
            row_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
            row_frame.grid(row=i, column=0, sticky="ew")
            row_frame.grid_columnconfigure(1, weight=1)

            cb = ctk.CTkCheckBox(
                row_frame, text="", variable=var,
                fg_color=SoP["ACCENT"], hover_color=SoP["ACCENT_HOVER"],
                width=30
            )
            cb.grid(row=0, column=0, padx=5, pady=(5, 0))

            # THIS BUTTON is for PREVIEW
            name_btn = ctk.CTkButton(
                row_frame,
                text=name,
                fg_color="transparent",
                text_color=SoP["TEXT"],
                hover_color=SoP["EDITOR"],
                anchor="w",
                command=lambda q=query: update_preview_panel(q)
            )
            name_btn.grid(row=0, column=1, sticky="ew", pady=(5, 0))

            # Store everything for filtering and importing
            self.extraction_vars.append((var, query))
            query_widgets.append((row_frame, name.lower(), query))

        # --- Search/Filter Function ---
        def filter_queries(*args):
            query = dialog_search_var.get().lower()
            for row_frame, name, query_dict in query_widgets:
                if query in name:
                    row_frame.grid()
                else:
                    row_frame.grid_remove()

        dialog_search_var.trace_add("write", filter_queries)

        # --- Action Buttons (Bottom) ---
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.grid(row=2, column=0, columnspan=2,
                       sticky="ew", padx=15, pady=10)
        btn_frame.grid_columnconfigure(3, weight=1)

        def select_all(val):
            for row_frame, name, query in query_widgets:
                if row_frame.winfo_viewable():  # Only affect visible items
                    # Find the var for this query
                    for var, q in self.extraction_vars:
                        if q['name'] == query['name']:
                            var.set(val)
                            break

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
            count = sum(1 for var, q in self.extraction_vars if var.get())
            import_btn.configure(text=f"Import {count} Queries")

        for var, q in self.extraction_vars:
            var.trace_add("write", update_btn_text)

    def _threaded_confirm_extraction(self, dialog: ctk.CTkToplevel):
        """Step 2: Gathers selected queries and passes them to the writer thread."""

        # This logic is now cleaner, just reads the list
        selected_queries = [query for var,
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
                        self.manager.store.root, "functions", safe_category)
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
