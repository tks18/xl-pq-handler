import customtkinter as ctk
from typing import Callable, List
import tkinter as tk
from tkinter import messagebox

from ....theme import SoP


class WorkbookExtractor(ctk.CTkFrame):
    def __init__(self, parent,
                 # Gets list from manager
                 refresh_list_callback: Callable[[], List[str]],
                 # Takes wb name, runs thread
                 extract_callback: Callable[[str], None],
                 **kwargs):
        super().__init__(
            parent, fg_color=SoP["FRAME"], corner_radius=8, **kwargs)

        self.refresh_list_callback = refresh_list_callback
        self.extract_callback = extract_callback
        self.open_workbook_names: List[str] = []

        self._build_widgets()
        self.refresh_workbook_list()  # Initial population

    def _build_widgets(self):
        self.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            self,
            text="Open Workbooks",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=SoP["ACCENT_HOVER"]
        )
        title.grid(row=0, column=0, sticky="ew")

        wb_frame = ctk.CTkFrame(self, fg_color="transparent")
        wb_frame.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 15))
        wb_frame.grid_columnconfigure(0, weight=1)

        self.workbook_menu = ctk.CTkOptionMenu(
            wb_frame, values=["Click 'Refresh'..."],
            fg_color=SoP["EDITOR"], button_color=SoP["ACCENT_DARK"],
            button_hover_color=SoP["ACCENT"], text_color=SoP["TEXT_DIM"],
            height=40
        )
        self.workbook_menu.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.refresh_wbs_btn = ctk.CTkButton(
            wb_frame, text="🔄", width=40, height=40, font=ctk.CTkFont(size=20),
            command=self.refresh_workbook_list,
            fg_color=SoP["TREE_FIELD"], hover_color=SoP["ACCENT_HOVER"]
        )
        self.refresh_wbs_btn.grid(row=0, column=1, padx=5)

        self.extract_btn = ctk.CTkButton(
            wb_frame, text="📥 Extract from Selected", height=40,
            command=self._on_extract,
            fg_color="transparent", border_width=1, border_color=SoP["ACCENT"],
            hover_color=SoP["TREE_FIELD"], text_color=SoP["ACCENT"], font=ctk.CTkFont(
                weight="bold")
        )
        self.extract_btn.grid(row=1, column=0, columnspan=2,
                              sticky="ew", pady=(10, 0))

    def refresh_workbook_list(self):
        """Gets workbook names via callback and updates the dropdown."""
        try:
            self.open_workbook_names = self.refresh_list_callback()
            if not self.open_workbook_names:
                self.workbook_menu.configure(
                    values=["No workbooks open."], state="disabled")
                self.workbook_menu.set("No workbooks open.")
                self.extract_btn.configure(state="disabled")
            else:
                self.workbook_menu.configure(
                    values=self.open_workbook_names, state="normal")
                self.workbook_menu.set(self.open_workbook_names[0])
                self.extract_btn.configure(state="normal")
        except Exception as e:
            messagebox.showerror(
                "Error", f"Failed to list open workbooks:\n{e}", parent=self)
            self.workbook_menu.configure(values=["Error"], state="disabled")
            self.extract_btn.configure(state="disabled")

    def _on_extract(self):
        """Calls the parent's extract thread function."""
        selected_name = self.workbook_menu.get()
        if not selected_name or selected_name in ["No workbooks open.", "Error", "Click 'Refresh'..."]:
            messagebox.showwarning(
                "No Workbook", "Please select a valid workbook.", parent=self)
            return
        self.extract_callback(selected_name)

    def set_busy(self, is_busy: bool):
        """Disables/Enables buttons during operation."""
        state = "disabled" if is_busy else "normal"
        text = "Reading..." if is_busy else "📥 Extract from Selected"
        self.extract_btn.configure(state=state, text=text)
        self.refresh_wbs_btn.configure(state=state)
        self.workbook_menu.configure(state=state)
