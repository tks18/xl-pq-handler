import customtkinter as ctk

from ...theme import BaseWidget


class PrimaryButton(BaseWidget, ctk.CTkButton):
    """Main action button (accent color)."""

    def __init__(self, master, style_manager, **kwargs):
        super().__init__(master, style_manager, self.style_widget, **kwargs)

    def style_widget(self):
        self.configure(
            fg_color=self.style_manager.theme.accent,
            hover_color=self.style_manager.theme.accent_hover,
            text_color=self.style_manager.theme.text,
            font=self.style_manager.font.heading,
            height=35,
            corner_radius=8
        )


class SecondaryButton(BaseWidget, ctk.CTkButton):
    """Subtle secondary action button."""

    def __init__(self, master, style_manager, **kwargs):
        super().__init__(master, style_manager, self.style_widget, **kwargs)

    def style_widget(self):
        self.configure(
            fg_color=self.style_manager.theme.frame,
            hover_color=self.style_manager.theme.accent_dark,
            text_color=self.style_manager.theme.text,
            font=self.style_manager.font.body,
            height=35,
            corner_radius=8
        )


class OutlinedButton(BaseWidget, ctk.CTkButton):
    """Transparent button with outline."""

    def __init__(self, master, style_manager, **kwargs):
        super().__init__(master, style_manager, self.style_widget, **kwargs)

    def style_widget(self):
        self.configure(
            fg_color="transparent",
            border_color=self.style_manager.theme.accent,
            border_width=2,
            text_color=self.style_manager.theme.accent,
            hover_color=self.style_manager.theme.accent_dark,
            corner_radius=8
        )


class TextButton(BaseWidget, ctk.CTkButton):
    """Minimal button — text only."""

    def __init__(self, master, style_manager, **kwargs):
        super().__init__(master, style_manager, self.style_widget, **kwargs)

    def style_widget(self):
        self.configure(
            fg_color="transparent",
            text_color=self.style_manager.theme.accent,
            hover_color=self.style_manager.theme.accent_dark
        )


class FloatingButton(PrimaryButton):
    """Circular floating button."""

    def style_widget(self):
        self.configure(
            width=50,
            height=50,
            corner_radius=25,
            fg_color=self.style_manager.theme.accent,
            hover_color=self.style_manager.theme.accent_hover
        )
