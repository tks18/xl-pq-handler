import customtkinter as ctk

from ...theme import BaseWidget


class TextField(BaseWidget, ctk.CTkEntry):
    def __init__(self, master, style_manager, placeholder="Enter text...", **kwargs):
        self.placeholder = placeholder
        super().__init__(master, style_manager, self.style_widget,
                         placeholder_text=placeholder, **kwargs)

    def style_widget(self):
        self.configure(
            fg_color=self.style_manager.theme.editor,
            text_color=self.style_manager.theme.text,
            border_color=self.style_manager.theme.accent_dark,
            corner_radius=6
        )


class SearchField(TextField):
    def style_widget(self):
        super().style_widget()
        self.configure(placeholder_text="Search...")


class PasswordField(TextField):
    def __init__(self, master, style_manager, **kwargs):
        super().__init__(master, style_manager, **kwargs)

    def style_widget(self):
        super().style_widget()
        self.configure(show="*")


class Dropdown(BaseWidget, ctk.CTkOptionMenu):
    def __init__(self, master, style_manager, values, **kwargs):
        super().__init__(master, style_manager, self.style_widget, values=values, **kwargs)

    def style_widget(self):
        self.configure(
            fg_color=self.style_manager.theme.frame,
            button_color=self.style_manager.theme.accent,
            text_color=self.style_manager.theme.text
        )
