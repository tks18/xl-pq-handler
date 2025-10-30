from .definitions import default_themes, BaseTheme
from .manager import StyleManager, BaseWidget

# Kept it only for the reason of compatibility with old code, Soon all the widgets will be refactored
# in future all the widgets should be based on BaseWidget and will auto support theme changing on the go
# instead of relying on the dictionary
SoP = {
    "BG": "#2d2b55",         # Darkest background
    "FRAME": "#282a4a",      # Sidebar, Panels
    "EDITOR": "#232540",     # Main content area bg
    "ACCENT": "#a277ff",     # Primary purple
    "ACCENT_HOVER": "#be99ff",  # Lighter purple for hover
    "ACCENT_DARK": "#4a468c",  # Darker purple for headers
    "TEXT": "#e0d9ef",       # Main text
    "TEXT_DIM": "#88849b",   # Dimmed text
    "TEXT_SUPER_DIM": "#535353",
    "TREE_FIELD": "#343261"  # Treeview row background,
}
