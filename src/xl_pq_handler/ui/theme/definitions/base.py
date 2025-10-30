from pydantic import BaseModel


class BaseTheme(BaseModel):
    """
    The Pydantic base model for a theme. Enforces that all
    fields are present and are strings.
    """
    name: str
    bg: str
    frame: str
    editor: str
    accent: str
    accent_hover: str
    accent_dark: str
    text: str
    text_dim: str
    tree_field: str
