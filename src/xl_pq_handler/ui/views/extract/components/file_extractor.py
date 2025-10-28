import customtkinter as ctk
from typing import Callable
import tkinter as tk
from tkinter import filedialog
import os

from ....theme import SoP


class FileExtractor(ctk.CTkFrame):
    def __init__(self, parent,
                 extract_callback: Callable[[str], None],
                 **kwargs):
        super().__init__(
            parent, fg_color=SoP["FRAME"], corner_radius=8, **kwargs)

        self.extract_callback = extract_callback
        self.selected_file_path = ""
        self._build_widgets()

    def _build_widgets(self):
        self.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            self,
            text="From Files",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=SoP["ACCENT_HOVER"]
        )
        title.grid(row=0, column=0, sticky="ew")

        file_frame = ctk.CTkFrame(self, fg_color="transparent")
        file_frame.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 15))

        file_frame.grid_columnconfigure(0, weight=0)
        file_frame.grid_columnconfigure(1, weight=1)

        self.select_file_btn = ctk.CTkButton(
            file_frame, text="📁 Select File...", height=40,
            command=self._select_excel_file,
            fg_color="transparent", border_width=1, border_color=SoP["ACCENT"],
            hover_color=SoP["TREE_FIELD"], text_color=SoP["ACCENT"]
        )
        self.select_file_btn.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.file_label = ctk.CTkLabel(
            file_frame, text="No file selected.", text_color=SoP["TEXT_DIM"],
            anchor="w",
            justify="left"
        )
        self.file_label.grid(row=0, column=1, sticky="ew")

        self.extract_btn = ctk.CTkButton(
            file_frame, text="Get Queries from File", height=40,
            command=self._on_extract,
            fg_color="transparent", border_width=1, border_color=SoP["ACCENT"],
            hover_color=SoP["TREE_FIELD"], text_color=SoP["ACCENT"], state="disabled"
        )
        self.extract_btn.grid(row=1, column=0, columnspan=2,
                              sticky="ew", pady=(10, 0))

    def _select_excel_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel Files", "*.xlsx;*.xlsm;*.xlsb"), ("All Files", "*.*")))
        if path:
            self.selected_file_path = path
            self.file_label.configure(text=os.path.basename(path))
            self.extract_btn.configure(
                state="normal", text_color=SoP["ACCENT"])
        else:
            self.selected_file_path = ""
            self.file_label.configure(text="No file selected.")
            self.extract_btn.configure(
                state="disabled", text_color=SoP["TEXT_DIM"])

    def _on_extract(self):
        if not self.selected_file_path:
            return
        self.extract_callback(self.selected_file_path)

    def set_busy(self, is_busy: bool):
        """Disables/Enables buttons during operation."""
        state = "disabled" if is_busy else "normal"
        text = "Reading..." if is_busy else "Get Queries from File"
        self.extract_btn.configure(state=state, text=text)
        self.select_file_btn.configure(state=state)
        if not is_busy and self.selected_file_path:  # Re-enable color if file selected
            self.extract_btn.configure(text_color=SoP["ACCENT"])
        elif not is_busy:
            self.extract_btn.configure(text_color=SoP["TEXT_DIM"])
