import tkinter as tk
import customtkinter as ctk
from typing import Callable, Dict, List, Tuple

from ....theme import SoP
from ....components import CTkCodeView
from .....classes import PQManager, PowerQueryScript


class BottomPanel(ctk.CTkFrame):
    """
    Component for the bottom panel in the Library view,
    containing the TabView and Action buttons/controls.
    """

    def __init__(self,
                 parent,
                 manager: PQManager,
                 insert_callback: Callable,
                 clear_selection_callback: Callable,
                 refresh_wb_list_callback: Callable,
                 format_tree_string_callback: Callable,  # Pass tree formatter
                 **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self.manager = manager
        self.insert_callback = insert_callback
        self.clear_selection_callback = clear_selection_callback
        self.refresh_wb_list_callback = refresh_wb_list_callback
        self.format_tree_string_callback = format_tree_string_callback

        self.grid_columnconfigure(0, weight=1)  # TabView area
        self.grid_columnconfigure(1, weight=0)  # Action Panel area
        self.grid_rowconfigure(0, weight=1)

        self.tab_widgets: Dict[str, ctk.CTkTextbox | CTkCodeView] = {}

        self._build_widgets()

    def _build_widgets(self):
        """Builds the TabView and Action Panel."""

        # --- TabView ---
        self.tabview = ctk.CTkTabview(
            self,
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
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self._create_tab_widgets(self.tabview)

        # --- Action Panel ---
        action_panel = ctk.CTkFrame(self, fg_color="transparent", width=180)
        action_panel.grid(row=0, column=1, sticky="nse", padx=5)
        action_panel.grid_columnconfigure(0, weight=1)

        # Workbook selector dropdown and refresh
        wb_select_frame = ctk.CTkFrame(action_panel, fg_color="transparent")
        wb_select_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        wb_select_frame.grid_columnconfigure(0, weight=1)

        self.insert_wb_menu = ctk.CTkOptionMenu(
            wb_select_frame, values=["Default (Active)"],
            fg_color=SoP["EDITOR"], button_color=SoP["ACCENT_DARK"],
            button_hover_color=SoP["ACCENT"], text_color=SoP["TEXT_DIM"],
            height=35, dynamic_resizing=False
        )
        self.insert_wb_menu.grid(row=0, column=0, sticky="ew")

        self.refresh_insert_wbs_btn = ctk.CTkButton(
            wb_select_frame, text="🔄", width=35, height=35,
            font=ctk.CTkFont(size=20),
            command=self.refresh_wb_list_callback,  # Call parent's method
            fg_color=SoP["TREE_FIELD"], hover_color=SoP["ACCENT_HOVER"]
        )
        self.refresh_insert_wbs_btn.grid(row=0, column=1, padx=(5, 0))

        # Main action buttons
        self.insert_btn = ctk.CTkButton(
            action_panel, text="➕ Insert Selected", height=40,
            command=self.insert_callback,  # Call parent's insert method
            fg_color=SoP["ACCENT"], hover_color=SoP["ACCENT_HOVER"],
            text_color="#000000", font=ctk.CTkFont(weight="bold")
        )
        self.insert_btn.grid(row=1, column=0, sticky="ew", pady=(5, 10))

        self.clear_btn = ctk.CTkButton(
            action_panel, text="Clear Selection", height=35,
            command=self.clear_selection_callback,  # Call parent's clear method
            fg_color="transparent", border_width=1,
            border_color=SoP["TEXT_DIM"], hover_color=SoP["TREE_FIELD"],
            text_color=SoP["TEXT_DIM"]
        )
        self.clear_btn.grid(row=2, column=0, sticky="ew", pady=5)

        self.selection_count_lbl = ctk.CTkLabel(
            action_panel, text="Selected: 0", text_color=SoP["TEXT_DIM"])
        self.selection_count_lbl.grid(row=3, column=0, sticky="ew", pady=10)

    def _create_tab_widgets(self, tabview):
        """Creates the text widgets inside the TabView."""
        self.tab_widgets["desc"] = ctk.CTkTextbox(
            tabview.add("Description"), fg_color="transparent", font=("Consolas", 12))
        self.tab_widgets["params"] = ctk.CTkTextbox(
            tabview.add("Parameters"), fg_color="transparent", font=("Consolas", 12))
        self.tab_widgets["deps"] = ctk.CTkTextbox(
            tabview.add("Dependencies"), fg_color="transparent", font=("Consolas", 12))
        self.tab_widgets["graph"] = ctk.CTkTextbox(
            tabview.add("Graph"), fg_color="transparent", font=("Consolas", 12))
        self.tab_widgets["sources"] = ctk.CTkTextbox(
            tabview.add("Data Sources"), fg_color="transparent", font=("Consolas", 12))
        self.tab_widgets["preview"] = CTkCodeView(
            tabview.add("Preview"), manager=self.manager)

        for name, widget in self.tab_widgets.items():
            widget.pack(fill="both", expand=True, padx=(
                0 if name == "preview" else 5), pady=(0 if name == "preview" else 5))
            widget.configure(state="disabled")
            self._configure_tab_tags(widget, name)
            # Add initial placeholder text
            placeholder = f"Select a query to view {name}."
            if name == "desc":
                placeholder = "Select query(s) to view description."
            elif name == "preview":
                placeholder = "Select a query to preview M-code."
            widget.insert("1.0", placeholder, ("dim",))
            widget.configure(state="disabled")  # Re-disable after insert

    def _configure_tab_tags(self, widget, tab_name):
        """Configures tags for a specific tab widget."""
        # Ensure dim tag exists for all
        widget.tag_config("dim", foreground=SoP["TEXT_DIM"])

        if isinstance(widget, CTkCodeView):
            # CodeView handles its own code tags, just ensure 'dim' is available
            pass
        else:
            # Configure tags for standard textboxes
            if tab_name == "params":
                widget.tag_config("name", foreground=SoP["ACCENT_HOVER"])
                widget.tag_config("type", foreground="#ce9178")
            elif tab_name == "sources":
                widget.tag_config("type", foreground="#c586c0")
                widget.tag_config("source", foreground="#ce9178")

    # --- Public Update Methods (Called by Parent View) ---

    def update_selection_count(self, count: int):
        """Updates the 'Selected: X' label."""
        self.selection_count_lbl.configure(text=f"Selected: {count}")

    def update_workbook_list(self, workbook_names: List[str]):
        """Updates the workbook dropdown for insertion."""
        menu_values = ["Default (Active)"] + workbook_names
        self.insert_wb_menu.configure(values=menu_values)
        self.insert_wb_menu.set("Default (Active)")
        if not workbook_names:
            self.insert_wb_menu.configure(text_color_disabled=SoP["TEXT_DIM"])

    def update_description(self, item_values_list: List[Tuple[str, ...]]):
        """Updates the Description tab based on selected items."""
        widget = self.tab_widgets["desc"]
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        if not item_values_list:
            widget.insert(
                "1.0", "Select row(s) to view description.", ("dim",))
        else:
            descs = []
            for i, vals in enumerate(item_values_list):
                if i >= 10:
                    descs.append(
                        f"\n... and {len(item_values_list) - 10} more ...")
                    break
                name, _, descr, _ = vals
                descs.append(f"--- {name} ---\n{descr or 'No description.'}")
            widget.insert("1.0", "\n\n".join(descs))
        widget.configure(state="disabled")

    def update_single_selection_tabs(self, script: PowerQueryScript | None):
        """Updates tabs that only show info for a single selected query."""
        # Clear all single-selection tabs first
        tabs_to_update = ["params", "deps", "graph", "sources", "preview"]
        for name in tabs_to_update:
            widget = self.tab_widgets[name]
            widget.configure(state="normal")
            widget.delete("1.0", "end")
        if not script:
            # Insert placeholder text if no script or multiple selected
            placeholder = "Select a single query to view {}."
            self.tab_widgets["params"].insert(
                "1.0", placeholder.format("parameters"), ("dim",))
            self.tab_widgets["deps"].insert(
                "1.0", placeholder.format("dependencies"), ("dim",))
            self.tab_widgets["graph"].insert(
                "1.0", placeholder.format("dependency graph"), ("dim",))
            self.tab_widgets["sources"].insert(
                "1.0", placeholder.format("data sources"), ("dim",))
            self.tab_widgets["preview"].insert(
                "1.0", placeholder.format("M-code"), ("dim",))
            # Re-disable after insert
            for name in tabs_to_update:
                self.tab_widgets[name].configure(state="disabled")
            return
        else:

            # --- Update Parameters ---
            try:
                widget = self.tab_widgets["params"]
                widget.configure(state="normal")
                params = self.manager.get_parameters_from_code(script.body)
                if not params:
                    widget.insert("1.0", "Not a function.", ("dim",))
                else:
                    for p in params:
                        widget.insert("end", f"Name:     ", ("dim",))
                        widget.insert("end", f"{p['name']}\n", ("name",))
                        widget.insert("end", f"Type:     ", ("dim",))
                        widget.insert("end", f"{p['type']}\n", ("type",))
                        widget.insert("end", f"Optional: ", ("dim",))
                        widget.insert(
                            "end", f"{'Yes' if p['optional'] else 'No'}\n\n", ("name",))
            except Exception as e:
                self._show_tab_error("params", f"Parse error: {e}")

            # --- Update Dependencies (List) ---
            try:
                widget = self.tab_widgets["deps"]
                widget.configure(state="normal")
                deps_list = script.meta.dependencies
                if deps_list:
                    widget.insert(
                        "1.0", f"Dependencies for {script.meta.name}:\n")
                    for d in deps_list:
                        widget.insert("end", f"\n • {d}")
                else:
                    widget.insert("1.0", "No dependencies.", ("dim",))
            except Exception as e:
                self._show_tab_error("deps", f"Error: {e}")

            # --- Update Dependency Graph ---
            try:
                widget = self.tab_widgets["graph"]
                widget.configure(state="normal")
                tree_data = self.manager.resolver.get_dependency_tree(
                    script.meta.name)
                tree_string = self.format_tree_string_callback(tree_data)
                widget.insert("1.0", tree_string)
            except Exception as e:
                self._show_tab_error("graph", f"Could not build graph: {e}")

            # --- Update Data Sources ---
            try:
                widget = self.tab_widgets["sources"]
                widget.configure(state="normal")
                sources = self.manager.get_datasources_from_code(script.body)
                if not sources:
                    widget.insert("1.0", "No external data sources.", ("dim",))
                else:
                    for src in sources:
                        widget.insert("end", f"Type:   ", ("dim",))
                        widget.insert("end", f"{src['type']}\n", ("type",))
                        widget.insert("end", f"Source: ", ("dim",))
                        widget.insert(
                            "end", f"{src['full_argument']}", ("source",))
                        suffix = ""
                        if src["source_type"] == "Variable":
                            suffix = " (Input Parameter)"
                        elif src["source_type"] == "Literal":
                            suffix = " (Literal)"
                        widget.insert("end", f"{suffix}\n\n", ("dim",))
            except Exception as e:
                self._show_tab_error(
                    "sources", f"Could not parse sources: {e}")

            # --- Update Preview ---
            try:
                code_widget = self.tab_widgets["preview"]
                if isinstance(code_widget, CTkCodeView):
                    code_widget.set_code(script.body)
            except Exception as e:
                self._show_tab_error("preview", f"Could not render code: {e}")

        for name in tabs_to_update:
            self.tab_widgets[name].configure(state="disabled")

    def _show_tab_error(self, tab_name: str, message: str):
        """Helper to display an error message in a tab."""
        widget = self.tab_widgets[tab_name]
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", message, ("dim",))
        widget.configure(state="disabled")

    # --- Public method for parent to get selected workbook ---
    def get_selected_workbook_target(self) -> str | None:
        """Returns the name selected in the insert dropdown, or None for default."""
        selected = self.insert_wb_menu.get()
        if selected == "Default (Active)":
            return None
        return selected
