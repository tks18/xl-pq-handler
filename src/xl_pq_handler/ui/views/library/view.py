# ui_library_view.py
import tkinter as tk
from tkinter import messagebox
from typing import Callable
import customtkinter as ctk
import threading
import pandas as pd

from ....classes import PQManager
from .components import EditMetadataDialog, TopBar, TreeviewArea, BottomPanel
from ...theme import SoP


class LibraryView(ctk.CTkFrame):
    """
    Manages the 'Library' view, including the query browser,
    search, filters, and info panels.
    """

    def __init__(self, parent, manager: PQManager, refresh_callback: Callable):
        super().__init__(parent, fg_color="transparent")
        self.manager = manager

        self.df = pd.DataFrame()
        self.categories = []
        self.cat_vars = {}
        self.sort_column = "Name"
        self.sort_asc = True
        self.open_workbooks_for_insert = []
        self.refresh_callback = refresh_callback

        self._init_vars()

        self._build_widgets()
        self._build_context_menu()
        self.refresh_data()  # Load initial data

    def _init_vars(self):
        """Initializes shared variables."""
        self.search_var = tk.StringVar()

    def _build_widgets(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=3)  # Treeview
        self.grid_rowconfigure(2, weight=1)  # Bottom Panel

        # --- Top Bar (Search + Categories) ---
        self._build_top_bar()

        # --- Tree Area ---
        self._build_treeview_area()

        # --- Bottom Panel (Description + Actions) ---
        self._build_bottom_panel()

        self.clear_selection()

    def _build_top_bar(self):
        """Builds the search entry and category filter section."""
        self.top_bar = TopBar(
            parent=self,
            search_var=self.search_var,
            search_callback=self._focus_first_result,  # Pass method for Enter key
            category_popup_callback=self._open_category_popup  # Pass method for button click
        )
        self.top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.search_var.trace_add("write", lambda *a: self.populate_tree())

    def _build_treeview_area(self):
        """Builds the Treeview widget and its container."""
        self.treeview_area = TreeviewArea(
            parent=self,
            select_callback=self._on_tree_select,  # Pass method for selection change
            sort_callback=self._sort_by_col,      # Pass method for sorting
            double_click_callback=self._on_double_click_insert,  # Pass method for double-click
            enter_callback=self._on_enter_insert  # Pass method for Enter key
        )
        self.treeview_area.grid(row=1, column=0, sticky="nsew")

    def _build_bottom_panel(self):
        """Builds the bottom TabView and Action Panel."""
        self.bottom_panel = BottomPanel(
            parent=self,
            manager=self.manager,
            insert_callback=self._threaded_insert_selected,  # Pass insert method
            clear_selection_callback=self.clear_selection,  # Pass clear method
            refresh_wb_list_callback=self._refresh_workbook_list_for_insert_wrapper,  # Pass wrapper
            format_tree_string_callback=self._format_tree_string  # Pass helper
        )
        self.bottom_panel.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        self._refresh_workbook_list_for_insert_wrapper()

    def _build_context_menu(self):
        """Creates the right-click context menu."""
        self.context_menu = tk.Menu(self.treeview_area.tree, tearoff=0)
        self.context_menu.add_command(
            label="Edit Metadata...",
            command=self._on_edit_metadata
        )
        self.context_menu.add_command(
            label="Open in VS Code / Editor",
            command=self._on_open_in_editor
        )
        self.context_menu.add_separator()
        self.context_menu.add_command(
            label="Delete", command=self._on_delete_query)

        # Bind to the Treeview
        self.treeview_area.tree.bind(
            "<Button-3>", self._show_context_menu)  # Windows/Linux
        self.treeview_area.tree.bind(
            "<Button-2>", self._show_context_menu)  # macOS

    def _get_selected_query_name(self) -> str | None:
        """Helper to get the name of the currently focused query."""
        return self.treeview_area.get_focused_item_name()

    def _on_open_in_editor(self):
        """Callback for 'Open in Editor' menu item."""
        name = self._get_selected_query_name()
        if not name:
            return

        try:
            self.manager.open_in_editor(name)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open editor:\n{e}")

    def _on_edit_metadata(self):
        """Callback for 'Edit Metadata' menu item."""
        name = self._get_selected_query_name()
        if not name:
            return

        script = self.manager.get_script(name)
        if not script:
            messagebox.showerror("Error", f"Could not find script '{name}'")
            return

        # --- Launch the Edit Dialog ---
        dialog = EditMetadataDialog(
            parent=self,
            manager=self.manager,
            script_to_edit=script,
            refresh_callback=self.refresh_callback
        )

    def _on_delete_query(self):
        """Callback for 'Delete' menu item."""
        name = self._get_selected_query_name()
        if not name:
            return

        if not messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{name}'?\n\nThis action cannot be undone."):
            return

        try:
            self.manager.store.delete_script(name)
            # We must tell the main app to refresh everything
            self.refresh_callback()
        except Exception as e:
            messagebox.showerror("Error", f"Could not delete query:\n{e}")

    def _show_context_menu(self, event):
        """Shows the context menu at the cursor's location."""
        iid = self.treeview_area.tree.identify_row(event.y)
        if iid:
            # Select the row under the cursor
            self.treeview_area.tree.selection_set(iid)
            self.treeview_area.tree.focus(iid)
            # Display the menu
            self.context_menu.post(event.x_root, event.y_root)

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
        summary_text = "All categories"
        button_text = "Categories ▾"

        if chosen and len(chosen) != len(self.categories):
            short = ", ".join(chosen[:3])
            if len(chosen) > 3:
                short += f", +{len(chosen)-3}"
            summary_text = short
            button_text = f"{len(chosen)} selected ▾"

        # Call methods on the TopBar instance
        self.top_bar.update_summary_text(summary_text)
        self.top_bar.update_button_text(button_text)

    def _apply_category_selection(self, popup):
        self._update_category_summary()
        if popup:
            popup.destroy()
        self.populate_tree()

    def populate_tree(self, *_):
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

        self.treeview_area.populate(dff)

        count = len(self.treeview_area.get_selection())
        self.bottom_panel.update_selection_count(count)
        self.clear_selection()

    # --- Selection & Sorting ---
    def _format_tree_string(self, node: dict, indent: str = "") -> str:
        """Recursively formats the dependency tree dict into a string."""
        tree_str = f"{indent}• {node['name']}\n"

        children = node.get("children", [])
        for i, child in enumerate(children):
            is_last = (i == len(children) - 1)
            child_indent = indent + ("    " if is_last else "│   ")
            tree_str += self._format_tree_string(child, child_indent)
        return tree_str

    def _on_tree_select(self, event=None):
        """Main handler for Treeview selection changes."""
        selected_iids = self.treeview_area.get_selection()
        self.bottom_panel.update_selection_count(len(selected_iids))

        # Update description panel (handles multi-select)
        item_values_list = [self.treeview_area.get_item_values(
            iid) for iid in selected_iids]
        self.bottom_panel.update_description(item_values_list)

        # Update panels that require single selection
        script = None
        if len(selected_iids) == 1:
            name = self.treeview_area.get_item_values(selected_iids[0])[0]
            script = self.manager.get_script(name)
            # Handle script not found if necessary (maybe inside update_single_selection_tabs)
        self.bottom_panel.update_single_selection_tabs(script)

    def clear_selection(self):
        self.treeview_area.clear_selection()
        self._on_tree_select()

    def _focus_first_result(self):
        self.treeview_area.focus_first_item()
        self._on_tree_select()

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

    def _on_double_click_insert(self, event):
        name = self._get_selected_query_name()
        if name:
            self._threaded_insert_selected(names=[name], single_only=True)

    def _threaded_insert_selected(self, names: list[str] | None = None, single_only=False):
        actual_names = []
        if names:  # If called directly with names (e.g., from double-click)
            actual_names = names
        else:  # If called from button click, get current selection
            selected_iids = self.treeview_area.get_selection()
            if not selected_iids:
                messagebox.showwarning(
                    "No selection", "Please select functions.")
                return
            actual_names = [self.treeview_area.get_item_values(
                iid)[0] for iid in selected_iids]

        if not actual_names:
            return
        workbook_name_arg = self.bottom_panel.get_selected_workbook_target()

        threading.Thread(
            target=self.insert_selected_functions,
            args=(actual_names, workbook_name_arg),
            daemon=True
        ).start()

    def _refresh_workbook_list_for_insert_wrapper(self):
        """Gets workbook names and tells BottomPanel to update."""
        try:
            names = self.manager.excel.list_open_workbooks()
            self.bottom_panel.update_workbook_list(names)
        except Exception as e:
            print(f"Could not refresh workbook list for insert: {e}")
            self.bottom_panel.update_workbook_list(
                [])  # Update with empty list on error

    def insert_selected_functions(self, names: list[str], workbook_name: str | None):
        """
        Thread target to insert queries.
        Uses the new manager API with try/except.
        """
        try:
            self.manager.insert_into_excel(
                names=names, workbook_name=workbook_name)

            target_desc = workbook_name if workbook_name else "the active workbook"
            summary = f"✅ Successfully Inserted {len(names)} Queries (and dependencies) into:\n{target_desc}"
            summary += "\n".join(f"  - {name}" for name in names)

            self.master.after(0, lambda: messagebox.showinfo(
                "Insertion Complete", summary))

        except Exception as e:
            self.master.after(0, lambda: messagebox.showerror(
                "Insertion Error", str(e)))
