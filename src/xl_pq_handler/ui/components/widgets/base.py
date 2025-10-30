from typing import Callable
import customtkinter as ctk

from ...theme import StyleManager, BaseWidget


class PrimaryButton(BaseWidget, ctk.CTkButton):
    """The main 'accent' button (e.g., Save, Insert, Extract)."""

    def __init__(self, master, style_manager: StyleManager, **kwargs):

        super().__init__(master, style_manager, self.style_widget, **kwargs)

    def style_widget(self):
        self.configure(fg_color=self.style_manager.theme.accent)
        self.configure(hover_color=self.style_manager.theme.accent_hover)
        self.configure(text_color="#000000")
        self.configure(font=self.style_manager.font.heading)
        self.configure(height=35)
