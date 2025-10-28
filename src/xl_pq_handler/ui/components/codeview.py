# ui_codeview.py
import tkinter as tk
import customtkinter as ctk
from ..theme import SoP
from ...classes import PQManager


class CTkCodeView(ctk.CTkFrame):
    """
    A CustomTkinter Frame with built-in M-code syntax highlighting.
    """

    def __init__(self, master, manager: PQManager, width=400, height=300, corner_radius=0, **kwargs):
        super().__init__(master, width=width, height=height, corner_radius=corner_radius)

        self.manager = manager

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)

        # --- Code Text Box ---
        # We need access to the underlying tk.Text for some bindings
        self.textbox = ctk.CTkTextbox(
            self,
            corner_radius=0,
            fg_color=SoP["EDITOR"],
            text_color=SoP["TEXT"],
            font=("Consolas", 12),
            wrap="none",
            **kwargs  # Pass other CTkTextbox options
        )
        self.textbox.grid(row=0, column=1, sticky="nsew")

        self._configure_syntax_tags()

    def _configure_syntax_tags(self):
        """Defines the color tags for the internal textbox."""
        self.textbox.tag_config("KEYWORD", foreground=SoP["ACCENT_HOVER"])
        self.textbox.tag_config("FUNCTION", foreground="#4ec9b0")  # Teal
        self.textbox.tag_config("DATASOURCE", foreground="#c586c0")  # Mauve
        self.textbox.tag_config("STRING", foreground="#ce9178")  # Orange
        self.textbox.tag_config("NUMBER", foreground="#b5cea8")  # Light Green
        self.textbox.tag_config("COMMENT", foreground=SoP["TEXT_DIM"])
        self.textbox.tag_config("OPERATOR", foreground=SoP["TEXT_DIM"])
        self.textbox.tag_config("IDENTIFIER", foreground=SoP["TEXT"])
        self.textbox.tag_config("OTHER", foreground="#d4d4d4")
        # tags needed by other tabs (params, sources) for consistency
        self.textbox.tag_config("dim", foreground=SoP["TEXT_DIM"])
        self.textbox.tag_config("name", foreground=SoP["ACCENT_HOVER"])
        self.textbox.tag_config("type", foreground="#ce9178")
        self.textbox.tag_config("source", foreground="#ce9178")

    def set_code(self, body: str):
        """Clears the textbox and inserts new code with syntax highlighting."""
        self.textbox.configure(state="normal")
        self.textbox.delete("1.0", "end")

        if not body:  # Handle empty input gracefully
            self.textbox.configure(state="disabled")
            return

        # Use the stored manager instance
        tokens = self.manager.get_tokens_from_code(body)

        for value, kind in tokens:
            start_index = self.textbox.index("end-1c")
            self.textbox.insert("end", value, (kind,))
            # Re-apply tag for multi-line tokens (important for comments/strings)
            if '\n' in value and kind in ["COMMENT", "STRING"]:
                end_index = self.textbox.index("end-1c")
                self.textbox.tag_add(kind, start_index, end_index)

        self.textbox.configure(state="disabled")

    # --- Methods to mimic CTkTextbox ---

    def insert(self, index, text, tags=None):
        self.textbox.insert(index, text, tags)

    def delete(self, start_index, end_index=None):
        self.textbox.delete(start_index, end_index)

    def configure(self, require_redraw=False, **kwargs):
        if 'state' in kwargs:
            self.textbox.configure(state=kwargs.pop('state'))
        super().configure(require_redraw, **kwargs)  # Configure the frame itself

    def get(self, start_index, end_index=None):
        return self.textbox.get(start_index, end_index)

    def index(self, index):
        return self.textbox.index(index)

    def tag_config(self, tagName, **kwargs):
        self.textbox.tag_config(tagName, **kwargs)

    def tag_add(self, tagName, start_index, end_index):
        self.textbox.tag_add(tagName, start_index, end_index)
