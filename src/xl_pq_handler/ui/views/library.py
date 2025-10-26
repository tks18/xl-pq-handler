# ui_library_view.py
import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
import threading
import pandas as pd
from ...classes import PQManager
from ..theme import SoP


class LibraryView(ctk.CTkFrame):
    """
    Manages the 'Library' view, including the query browser,
    search, filters, and info panels.
    """

    def __init__(self, parent, manager: PQManager):
        super().__init__(parent, fg_color="transparent")
        self.manager = manager

        self.df = pd.DataFrame()
        self.categories = []
        self.cat_vars = {}
        self.sort_column = "Name"
        self.sort_asc = True

        self._build_widgets()
        self.refresh_data()  # Load initial data

    def _build_widgets(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=3)  # Treeview
        self.grid_rowconfigure(2, weight=1)  # Bottom Panel

        # --- Top Bar (Search + Categories) ---
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        top.grid_columnconfigure(0, weight=1)

        self.search_var = tk.StringVar()
        self.search_entry = ctk.CTkEntry(
            top, textvariable=self.search_var,
            placeholder_text="Filter by name, category, description, or tags...",
            fg_color=SoP["TREE_FIELD"], border_color=SoP["FRAME"],
            text_color=SoP["TEXT"], height=35)
        self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.search_entry.bind(
            "<Return>", lambda e: self._focus_first_result())
        self.search_var.trace_add("write", lambda *a: self.populate_tree())

        self.cat_button = ctk.CTkButton(
            top, text="Categories ▾", width=160, height=35,
            command=self._open_category_popup, fg_color=SoP["FRAME"],
            border_color=SoP["ACCENT_DARK"], border_width=1,
            hover_color=SoP["TREE_FIELD"], text_color=SoP["TEXT_DIM"])
        self.cat_button.grid(row=0, column=1, padx=6)

        self.cat_summary = ctk.CTkLabel(
            top, text="All categories", text_color=SoP["TEXT_DIM"])
        self.cat_summary.grid(row=0, column=2, padx=(10, 5), sticky="w")

        # --- Tree Area ---
        tree_container = ctk.CTkFrame(
            self, fg_color=SoP["TREE_FIELD"], corner_radius=8)
        tree_container.grid(row=1, column=0, sticky="nsew")
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        self.columns = ("Name", "Category", "Description", "Version")
        self.tree = ttk.Treeview(
            tree_container, columns=self.columns, show="headings", selectmode="extended")

        self.tree.column("Name", width=220, anchor="w")
        self.tree.column("Category", width=140, anchor="w")
        self.tree.column("Description", width=460, anchor="w")
        self.tree.column("Version", width=80, anchor="center")
        for col in self.columns:
            self.tree.heading(
                col, text=col, command=lambda c=col: self._sort_by_col(c))

        vsb = ttk.Scrollbar(tree_container, orient="vertical",
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both",
                       expand=True, padx=(1, 0), pady=(1, 0))

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except:
            pass
        style.configure(
            "Treeview", background=SoP["TREE_FIELD"],
            fieldbackground=SoP["TREE_FIELD"], foreground=SoP["TEXT"], rowheight=28)
        style.configure(
            "Treeview.Heading", background=SoP["ACCENT_DARK"],
            foreground=SoP["ACCENT_HOVER"], relief="flat", font=("Calibri", 11, "bold"))
        style.map("Treeview.Heading", background=[
                  ("active", SoP["ACCENT_DARK"])])
        style.map("Treeview", background=[("selected", SoP["ACCENT"])])

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_double_click_insert)
        self.tree.bind("<Return>", self._on_enter_insert)
        self.tree.bind("<Shift-Down>", self._on_shift_select_down)
        self.tree.bind("<Shift-Up>", self._on_shift_select_up)

        # Bind Ctrl+A
        self.tree.bind("<Control-a>", self._select_all_visible)
        self.tree.bind("<Control-A>", self._select_all_visible)

        # --- Bottom Panel (Description + Actions) ---
        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_rowconfigure(0, weight=1)

        # --- TabView for Description/Dependencies/Preview ---
        self.bottom_tabview = ctk.CTkTabview(
            bottom_frame,
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
        self.bottom_tabview.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.bottom_tabview.add("Description")
        self.bottom_tabview.add("Dependencies")
        self.bottom_tabview.add("Preview")

        self.desc = ctk.CTkTextbox(self.bottom_tabview.tab(
            "Description"), fg_color="transparent", text_color=SoP["TEXT_DIM"])
        self.desc.pack(fill="both", expand=True, padx=5, pady=5)
        self.desc.configure(state="disabled")

        self.deps = ctk.CTkTextbox(self.bottom_tabview.tab(
            "Dependencies"), fg_color="transparent", text_color=SoP["TEXT_DIM"])
        self.deps.pack(fill="both", expand=True, padx=5, pady=5)
        self.deps.configure(state="disabled")

        self.preview = ctk.CTkTextbox(
            self.bottom_tabview.tab("Preview"),
            fg_color="transparent",
            text_color=SoP["TEXT"],
            font=("Consolas", 12)  # Monospaced font
        )
        self.preview.pack(fill="both", expand=True, padx=5, pady=5)
        self.preview.configure(state="disabled")

        self.desc.tag_config("dim", foreground=SoP["TEXT_DIM"])
        self.deps.tag_config("dim", foreground=SoP["TEXT_DIM"])
        self.preview.tag_config("dim", foreground=SoP["TEXT_DIM"])

        # --- Action Buttons ---
        action_panel = ctk.CTkFrame(
            bottom_frame, fg_color="transparent", width=180)
        action_panel.grid(row=0, column=1, sticky="ne", padx=5)
        self.insert_btn = ctk.CTkButton(
            action_panel, text="➕ Insert Selected", height=40,
            command=self._threaded_insert_selected, fg_color=SoP["ACCENT"],
            hover_color=SoP["ACCENT_HOVER"], text_color="#000000",
            font=ctk.CTkFont(weight="bold"))
        self.insert_btn.pack(fill="x", pady=(0, 10))
        self.clear_btn = ctk.CTkButton(
            action_panel, text="Clear Selection", height=35,
            command=self.clear_selection, fg_color="transparent", border_width=1,
            border_color=SoP["TEXT_DIM"], hover_color=SoP["TREE_FIELD"],
            text_color=SoP["TEXT_DIM"])
        self.clear_btn.pack(fill="x", pady=5)
        self.selection_count_lbl = ctk.CTkLabel(
            action_panel, text="Selected: 0", text_color=SoP["TEXT_DIM"])
        self.selection_count_lbl.pack(fill="x", pady=10)
        self.clear_selection()

    def _ensure_df_columns(self):
        """Ensure dataframe has all expected columns after loading."""
        if "tags" not in self.df.columns:
            self.df["tags"] = [[] for _ in range(len(self.df))]
        if "dependencies" not in self.df.columns:
            self.df["dependencies"] = [[] for _ in range(len(self.df))]

    def refresh_data(self):
        """Called by the main app to reload all data from the manager."""
        try:
            # Use the manager's store to get the dataframe
            self.df = self.manager.store.index_to_dataframe()
            self._ensure_df_columns()
            self.df["category"] = self.df["category"].fillna("Uncategorized")
            self.categories = sorted(self.df["category"].unique().tolist())

            # Preserve existing category selections
            new_cat_vars = {}
            for c in self.categories:
                if c in self.cat_vars:
                    new_cat_vars[c] = self.cat_vars[c]
                else:
                    new_cat_vars[c] = tk.BooleanVar(value=True)
            self.cat_vars = new_cat_vars

            self.populate_tree()
            self._update_category_summary()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load index:\n{e}")
            self.df = pd.DataFrame()  # Clear df on error

    def _open_category_popup(self):
        if hasattr(self, "_cat_popup") and self._cat_popup.winfo_exists():
            self._cat_popup.lift()
            return
        popup = ctk.CTkToplevel(self)
        popup.title("Select Categories")
        popup.geometry("420x440")
        popup.transient()
        popup.grab_set()
        popup.configure(fg_color=SoP["BG"])
        self._cat_popup = popup
        frame = ctk.CTkScrollableFrame(popup, fg_color=SoP["FRAME"])
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        for c in self.categories:
            var = self.cat_vars.get(c)
            cb = ctk.CTkCheckBox(
                frame, text=c, variable=var, fg_color=SoP["ACCENT"],
                hover_color=SoP["ACCENT_HOVER"], text_color=SoP["TEXT"])
            cb.pack(anchor="w", pady=6, padx=4)

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        apply_btn = ctk.CTkButton(
            btn_frame, text="Apply",
            command=lambda: self._apply_category_selection(popup),
            fg_color=SoP["ACCENT"], hover_color=SoP["ACCENT_HOVER"],
            text_color="#000000", font=ctk.CTkFont(weight="bold"))
        apply_btn.pack(side="right", padx=(6, 0))
        ctk.CTkButton(
            btn_frame, text="All", width=80,
            command=self._select_all_categories,
            fg_color=SoP["FRAME"], text_color=SoP["TEXT_DIM"]
        ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(
            btn_frame, text="None", width=80,
            command=self._clear_all_categories,
            fg_color=SoP["FRAME"], text_color=SoP["TEXT_DIM"]
        ).pack(side="left", padx=(6, 6))

    def _select_all_categories(self): [v.set(True)
                                       for v in self.cat_vars.values()]

    def _clear_all_categories(self): [v.set(False)
                                      for v in self.cat_vars.values()]

    def _update_category_summary(self):
        chosen = [c for c, v in self.cat_vars.items() if v.get()]
        if not chosen or len(chosen) == len(self.categories):
            self.cat_summary.configure(text="All categories")
            self.cat_button.configure(text="Categories ▾")
        else:
            short = ", ".join(chosen[:3])
            if len(chosen) > 3:
                short += f", +{len(chosen)-3}"
            self.cat_summary.configure(text=short)
            self.cat_button.configure(text=f"{len(chosen)} selected ▾")

    def _apply_category_selection(self, popup):
        self._update_category_summary()
        if popup:
            popup.destroy()
        self.populate_tree()

    def populate_tree(self, *_):
        for i in self.tree.get_children():
            self.tree.delete(i)

        if self.df.empty:
            return

        q = (self.search_var.get() or "").strip().lower()
        chosen = set([c for c, v in self.cat_vars.items() if v.get()])
        dff = self.df.copy()

        if chosen and len(chosen) != len(self.categories):
            dff = dff[dff["category"].isin(chosen)]

        if q:
            def _row_matches(r):
                if q in str(r["name"]).lower():
                    return True
                if q in str(r["category"]).lower():
                    return True
                if q in str(r["description"]).lower():
                    return True
                tags = r.get("tags")
                if isinstance(tags, list):
                    if any(q in str(tag).lower() for tag in tags):
                        return True
                return False
            dff = dff[dff.apply(_row_matches, axis=1)]

        col_map = {"Name": "name", "Category": "category",
                   "Description": "description", "Version": "version"}

        sort_col_df = col_map.get(self.sort_column, "name")
        if sort_col_df not in dff.columns:
            sort_col_df = "name"

        dff = dff.sort_values(by=sort_col_df, ascending=self.sort_asc)

        for _, row in dff.iterrows():
            ver = row.get("version", "")
            # Use name as IID, fallback to path if names collide
            iid = row["name"]
            if self.tree.exists(iid):
                iid = row["path"]

            self.tree.insert("", "end", iid=iid, values=(
                row["name"], row["category"], row.get("description", ""), ver))

        self.selection_count_lbl.configure(
            text=f"Selected: {len(self.tree.selection())}")
        self.clear_selection()

    # --- Selection & Sorting ---
    def _on_tree_select(self, event=None):
        sels = self.tree.selection()
        self.selection_count_lbl.configure(text=f"Selected: {len(sels)}")

        # --- Update Description Panel ---
        self.desc.configure(state="normal")
        self.desc.delete("1.0", "end")
        if not sels:
            self.desc.insert(
                "1.0", "Select one or more rows to view description(s).", ("dim",))
        else:
            descs = []
            for i, iid in enumerate(sels):
                if i >= 10:
                    descs.append(
                        f"\n... and {len(sels) - 10} more selected ...")
                    break
                vals = self.tree.item(iid, "values")
                name = vals[0]
                descr = vals[2] or "No description."
                descs.append(f"--- {name} ---\n{descr}")
            self.desc.insert("1.0", "\n\n".join(descs))
        self.desc.configure(state="disabled")

        # --- Update Dependencies/Preview (single selection only) ---
        self.deps.configure(state="normal")
        self.deps.delete("1.0", "end")
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")

        if len(sels) == 1:
            name = self.tree.item(sels[0], "values")[0]
            # NEW API CALL: Get the full script object
            script = self.manager.get_script(name)

            if script:
                # Update Dependencies
                deps_list = script.meta.dependencies
                if deps_list:
                    self.deps.insert("1.0", f"Dependencies for {name}:\n")
                    for d in deps_list:
                        self.deps.insert("end", f"\n • {d}")
                else:
                    self.deps.insert(
                        "1.0", f"{name} has no dependencies.", ("dim",))

                # Update Preview
                self.preview.insert("1.0", script.body)
            else:
                self.deps.insert(
                    "1.0", "Could not find query in index.", ("dim",))
                self.preview.insert(
                    "1.0", f"Error: Could not find or read file for {name}.", ("dim",))

        elif len(sels) > 1:
            self.deps.insert(
                "1.0", "Select a single query to view dependencies.", ("dim",))
            self.preview.insert(
                "1.0", "Select a single query to preview its M code.", ("dim",))
        else:
            self.deps.insert(
                "1.0", "Select a query to view its dependencies.", ("dim",))
            self.preview.insert(
                "1.0", "Select a query to preview its M code.", ("dim",))

        self.deps.configure(state="disabled")
        self.preview.configure(state="disabled")

    def clear_selection(self):
        for sel in self.tree.selection():
            self.tree.selection_remove(sel)
        self._on_tree_select(None)

    def _select_all_visible(self, event=None):
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children)
            self._on_tree_select(None)
        return "break"

    def _focus_first_result(self):
        children = self.tree.get_children()
        if children:
            self.tree.focus(children[0])
            self.tree.selection_set(children[0])
            self.tree.see(children[0])
            self._on_tree_select(None)

    def _sort_by_col(self, col_name):
        if self.sort_column == col_name:
            self.sort_asc = not self.sort_asc
        else:
            self.sort_column = col_name
            self.sort_asc = True
        self.populate_tree()

    # --- Key Handlers & Insert ---

    def _on_enter_insert(self, event):
        self._threaded_insert_selected()
        return "break"

    def _on_shift_select_down(self, event):
        focus = self.tree.focus()
        if not focus:
            return "break"
        next_item = self.tree.next(focus)
        if next_item:
            self.tree.selection_add(next_item)
            self.tree.focus(next_item)
            self.tree.see(next_item)
        self._on_tree_select()
        return "break"

    def _on_shift_select_up(self, event):
        focus = self.tree.focus()
        if not focus:
            return "break"
        prev_item = self.tree.prev(focus)
        if prev_item:
            self.tree.selection_add(prev_item)
            self.tree.focus(prev_item)
            self.tree.see(prev_item)
        self._on_tree_select()
        return "break"

    def _on_double_click_insert(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            self._threaded_insert_selected(single_only=True)

    def _threaded_insert_selected(self, single_only=False):
        sels = self.tree.selection()
        if not sels:
            messagebox.showwarning(
                "No selection", "Please select functions to insert.")
            return

        if single_only and len(sels) > 1:
            # Double-click shouldn't insert all, just the one clicked
            # The event already set the selection to one, so this is a safeguard
            return

        names = [self.tree.item(iid, "values")[0] for iid in sels]

        threading.Thread(
            target=self.insert_selected_functions,
            args=(names,),
            daemon=True
        ).start()

    def insert_selected_functions(self, names: list[str]):
        """
        Thread target to insert queries.
        Uses the new manager API with try/except.
        """
        try:
            # NEW API CALL: Simple, clean, and handles dependencies.
            self.manager.insert_into_excel(names=names)

            summary = f"✅ Successfully Inserted {len(names)} Queries (and their dependencies):\n"
            summary += "\n".join(f"  - {name}" for name in names)

            self.master.after(0, lambda: messagebox.showinfo(
                "Insertion Complete", summary))

        except Exception as e:
            # The new manager raises an error on failure, which is cleaner.
            self.master.after(0, lambda: messagebox.showerror(
                "Insertion Error", str(e)))
