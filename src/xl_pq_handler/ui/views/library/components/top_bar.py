import tkinter as tk
import customtkinter as ctk
from typing import Callable

from ....theme import SoP


class TopBar(ctk.CTkFrame):
    """
    Component for the top bar in the Library view,
    containing search and category filtering.
    """

    def __init__(self,
                 parent,
                 search_var: tk.StringVar,  # Receive search_var from parent
                 search_callback: Callable,
                 category_popup_callback: Callable,
                 **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)

        self.search_callback = search_callback
        self.category_popup_callback = category_popup_callback
        self.search_var = search_var  # Store the shared variable

        self.grid_columnconfigure(0, weight=1)  # Search entry expands

        self._build_widgets()

    def _build_widgets(self):
        # --- Search Entry ---
        self.search_entry = ctk.CTkEntry(
            self,
            textvariable=self.search_var,  # Use the passed-in variable
            placeholder_text="Filter by name, category, description, tags...",
            fg_color=SoP["TREE_FIELD"], border_color=SoP["FRAME"],
            text_color=SoP["TEXT"], height=35, font=ctk.CTkFont(family="Segoe UI")
        )
        self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        # Binding remains, but calls the passed callback indirectly via search_var trace
        # Optional: trigger search on Enter
        self.search_entry.bind("<Return>", lambda e: self.search_callback())

        # --- Category Button ---
        self.cat_button = ctk.CTkButton(
            self,
            text="Categories ▾", width=160, height=35,
            command=self.category_popup_callback,  # Call the passed-in function
            fg_color=SoP["FRAME"],
            border_color=SoP["ACCENT_DARK"], border_width=1,
            hover_color=SoP["TREE_FIELD"], text_color=SoP["TEXT_DIM"],
            font=ctk.CTkFont(family="Segoe UI")
        )
        self.cat_button.grid(row=0, column=1, padx=6)

        # --- Category Summary Label ---
        self.cat_summary_label = ctk.CTkLabel(
            self,
            text="All categories",
            text_color=SoP["TEXT_DIM"],
            font=ctk.CTkFont(family="Segoe UI"),
        )
        self.cat_summary_label.grid(row=0, column=2, padx=(10, 5), sticky="w")

    # --- Method for parent to update summary ---
    def update_summary_text(self, text: str):
        self.cat_summary_label.configure(text=text)

    # --- Method for parent to update button text ---
    def update_button_text(self, text: str):
        self.cat_button.configure(text=text)
