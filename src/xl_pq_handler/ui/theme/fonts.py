from pydantic import BaseModel
import customtkinter as ctk


class BaseFontSet(BaseModel):
    """
    Pydantic base model for a font set.
    Enforces that all required font styles are defined.
    """
    title: ctk.CTkFont
    heading: ctk.CTkFont
    body: ctk.CTkFont
    code: ctk.CTkFont
    small_dim: ctk.CTkFont

    class Config:
        arbitrary_types_allowed = True


class FontManager:
    """
    Manages loading and accessing font set objects.
    Access the active set via the .font property.
    e.g., font_manager.font.title
    """

    def __init__(self, default_set="Default"):

        DefaultFontSet = BaseFontSet(
            title=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
            heading=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            body=ctk.CTkFont(family="Segoe UI", size=13),
            code=ctk.CTkFont(family="Consolas", size=12),
            small_dim=ctk.CTkFont(family="Segoe UI", size=11, slant="italic")
        )
        self._font_sets = {
            "Default": DefaultFontSet,
        }

        # 'font' holds the active font set object
        self.font: BaseFontSet = self._font_sets.get(
            default_set, DefaultFontSet)

    def set_font_set(self, set_name: str):
        """Sets the active font set."""
        if set_name in self._font_sets:
            self.font = self._font_sets[set_name]
            print(f"Font set to {set_name}. Restart app to see changes.")
        else:
            print(f"Error: Font set '{set_name}' not found.")
