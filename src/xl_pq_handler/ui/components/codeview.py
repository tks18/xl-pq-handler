# ui_codeview.py
import tkinter as tk
import customtkinter as ctk
from ..theme import SoP
from ...classes import PQManager


class CTkCodeView(ctk.CTkFrame):
    """
    A CustomTkinter Frame with synchronized line numbers and
    built-in M-code syntax highlighting.
    """

    def __init__(self, master, manager: PQManager, width=400, height=300, corner_radius=0, **kwargs):
        super().__init__(master, width=width, height=height, corner_radius=corner_radius)

        self.manager = manager

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)

        # --- Line Numbers Canvas ---
        self.line_numbers = tk.Canvas(
            self,
            width=40,
            bg=SoP["EDITOR"],  # Match editor background
            highlightthickness=0,
            bd=0  # No border
        )
        self.line_numbers.grid(row=0, column=0, sticky="ns")

        # --- Code Text Box ---
        # We need access to the underlying tk.Text for some bindings
        self.textbox = ctk.CTkTextbox(
            self,
            corner_radius=0,
            fg_color=SoP["EDITOR"],
            text_color=SoP["TEXT"],
            font=("Consolas", 12),
            wrap="none",  # Important for line numbers
            yscrollcommand=self._on_text_scroll,
            **kwargs  # Pass other CTkTextbox options
        )
        self.textbox.grid(row=0, column=1, sticky="nsew")

        self._configure_syntax_tags()

        # --- Synchronization Logic ---
        # Bind events to update line numbers
        self.textbox.bind("<Configure>", self._redraw_line_numbers)
        # Update on text change
        self.textbox.bind("<KeyRelease>", self._redraw_line_numbers)

        # Initial drawing
        self._redraw_line_numbers()

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

    def _on_text_scroll(self, first, last):
        """Called when the textbox's internal scrollbar moves or content changes."""

        # --- ADD THIS CHECK ---
        # Make sure the textbox attribute exists before proceeding
        if not hasattr(self, 'textbox'):
            return  # Object not fully initialized yet
        # --- END CHECK ---

        # The textbox handles its own scrollbar display via .set() internally.
        # We ONLY need to redraw the line numbers based on the new view.
        self._redraw_line_numbers()
        # Also, manually scroll the canvas to match the text's top fraction
        self.line_numbers.yview_moveto(first)

    def _redraw_line_numbers(self, event=None):
        """Redraws the line numbers in the canvas."""
        if not hasattr(self, 'textbox') or not self.textbox.winfo_exists() or not self.textbox.winfo_viewable():
            return  # Widget not ready, skip redraw

        self.line_numbers.delete("all")

        # Get the y-fraction of the view's top edge (e.g., 0.0 means scrolled to top)
        # first_frac = float(self.textbox.yview()[0]) # Alternative way

        # Get the index of the first character visible at the top-left
        first_visible_index = self.textbox.index("@0,0")
        first_line = int(first_visible_index.split('.')[0])

        # Calculate the y-offset (in pixels) of the very first line ("1.0")
        # from the top of the entire text content.
        # If the text is scrolled, this might be negative.
        first_line_abs_bbox = self.textbox.dlineinfo("1.0")
        if first_line_abs_bbox is None:
            return  # No text content yet

        # The y-coordinate of the top edge of the visible text area,
        # relative to the top of the *entire* text content.
        visible_top_y = first_line_abs_bbox[1]

        # Estimate last visible line (can be slightly off but okay)
        last_line_approx = first_line + \
            int(self.textbox.winfo_height() / first_line_abs_bbox[3]) + 2
        total_lines = int(self.textbox.index("end-1c").split('.')[0])

        for i in range(first_line, min(last_line_approx, total_lines + 1)):
            line_bbox = self.textbox.dlineinfo(f"{i}.0")
            if line_bbox:
                x_pos = 38  # Right-align
                # line_bbox[1] is the absolute y-position of the line's top
                # Calculate the line's position *relative* to the visible top edge
                y_on_canvas = line_bbox[1] - visible_top_y

                # Draw if it's within the canvas height
                if 0 <= y_on_canvas <= self.line_numbers.winfo_height() + 10:  # Add buffer
                    self.line_numbers.create_text(
                        x_pos,
                        # Center vertically approx
                        y_on_canvas + line_bbox[4] / 2,
                        anchor="ne",
                        text=str(i),
                        fill=SoP["TEXT_SUPER_DIM"],
                        font=("Consolas", 11)
                    )

        # Sync canvas scroll using the fraction provided by the textbox itself
        self.line_numbers.yview_moveto(self.textbox.yview()[0])

    def set_code(self, body: str):
        """Clears the textbox and inserts new code with syntax highlighting."""
        self.textbox.configure(state="normal")
        self.textbox.delete("1.0", "end")

        if not body:  # Handle empty input gracefully
            self.textbox.configure(state="disabled")
            self._redraw_line_numbers()
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
        self._redraw_line_numbers()  # Update line numbers after setting code

    # --- Methods to mimic CTkTextbox ---

    def insert(self, index, text, tags=None):
        self.textbox.insert(index, text, tags)
        self._redraw_line_numbers()

    def delete(self, start_index, end_index=None):
        self.textbox.delete(start_index, end_index)
        self._redraw_line_numbers()

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
