from typing import Callable
from weakref import WeakSet
import customtkinter as ctk

from .fonts import FontManager

from .definitions.base import BaseTheme


class StyleManager:
    """
    Manages loading and propagating Styling changes to live widgets.
    Handles both theme changes and font set changes globally
    """

    def __init__(self, themes: list[BaseTheme], default_theme: str):

        self._themes = {theme.name: theme for theme in themes}
        self.theme_names = list(self._themes.keys())

        self.font_manager = FontManager()
        self.font = self.font_manager.font

        self.theme: BaseTheme = self._themes.get(
            default_theme, self._themes[self.theme_names[0]])

        self._subscribers: WeakSet[BaseWidget] = WeakSet()

    def set_theme(self, theme_name: str):
        """
        Sets the active theme.
        """
        if theme_name in self._themes:
            self.theme = self._themes[theme_name]
            print(f"Theme set to {theme_name}.")
        else:
            print(f"Error: Theme '{theme_name}' not found.")
        self._notify_subscribers()

    def set_font_set(self, set_name: str):
        """Sets the active font set."""
        self.font_manager.set_font_set(set_name)
        self.font = self.font_manager.font
        self._notify_subscribers()

    def register(self, widget):
        """
        Subscribes a widget to theme updates.
        The widget *must* have a method named _on_theme_update()
        """
        self._subscribers.add(widget)

    def _notify_subscribers(self):
        """
        Calls the _on_theme_update() method on all live widgets.
        """
        for widget in self._subscribers:
            if widget and hasattr(widget, "_on_theme_update"):
                widget._on_theme_update()


class BaseWidget(ctk.CTkBaseClass):
    def __init__(self, master, style_manager: StyleManager, apply_theming_callback: Callable[[], None], **kwargs):
        """
        A base class for customtkinter widgets
        that can be subscribed to theme updates
        with the App's ThemeManager

        - Handles theme updates without a restart
        """
        self.style_manager = style_manager
        self.apply_theming_callback = apply_theming_callback

        super().__init__(master, **kwargs)
        self.apply_theming_callback()

        self.style_manager.register(self)

    def _on_theme_update(self):
        """
        Required by the App's ThemeManager to update theming 
        changes globally without the need for a restart
        """
        self.apply_theming_callback()
