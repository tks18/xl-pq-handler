from .base import BaseTheme


SoPTheme = BaseTheme(
    name="Shades of Purple",
    bg="#2d2b55",
    frame="#282a4a",
    editor="#232540",
    accent="#a277ff",
    accent_hover="#be99ff",
    accent_dark="#4a468c",
    text="#e0d9ef",
    text_dim="#88849b",
    tree_field="#343261"
)


OceanicTheme = BaseTheme(
    name="Oceanic",
    bg="#0d1117",
    frame="#161b22",
    editor="#010409",
    accent="#58a6ff",
    accent_hover="#79c0ff",
    accent_dark="#2f6cab",
    text="#c9d1d9",
    text_dim="#8b949e",
    tree_field="#161b22"
)


themes = [SoPTheme, OceanicTheme]
