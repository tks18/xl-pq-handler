import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
from typing import Callable, List, Tuple
import pandas as pd

from ....theme import SoP


class TreeviewArea(ctk.CTkFrame):
    """
    Component for the Treeview area in the Library view.
    Manages the tree, scrollbar, styling, and basic events.
    """

    def __init__(self,
                 parent,
                 select_callback: Callable,   # Notify parent on selection
                 sort_callback: Callable,     # Notify parent on header click
                 double_click_callback: Callable,  # Notify parent on double-click
                 enter_callback: Callable,    # Notify parent on Enter key
                 **kwargs):
        super().__init__(
            parent, fg_color=SoP["TREE_FIELD"], corner_radius=8, **kwargs)

        self.select_callback = select_callback
        self.sort_callback = sort_callback
        self.double_click_callback = double_click_callback
        self.enter_callback = enter_callback

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_treeview()
        self._apply_style()

    def _build_treeview(self):
        """Creates the ttk.Treeview and Scrollbar."""
        self.columns = ("Name", "Category", "Description", "Version")
        self.tree = ttk.Treeview(
            self, columns=self.columns, show="headings", selectmode="extended")

        # Column definitions
        self.tree.column("Name", width=220, anchor="w")
        self.tree.column("Category", width=140, anchor="w")
        self.tree.column("Description", width=460, anchor="w")
        self.tree.column("Version", width=80, anchor="center")
        for col in self.columns:
            # Call the parent's sort method via callback
            self.tree.heading(
                col, text=col, command=lambda c=col: self.sort_callback(c))

        # Scrollbar
        vsb = ctk.CTkScrollbar(self, orientation="vertical",
                               command=self.tree.yview, width=15, bg_color=SoP["TREE_FIELD"], fg_color=SoP["EDITOR"], corner_radius=8)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")  # Pack scrollbar first
        self.tree.pack(side="left", fill="both", expand=True)  # Then pack tree

        # Bindings - call parent callbacks
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.select_callback())
        self.tree.bind("<Double-1>", lambda e: self.double_click_callback(e))
        self.tree.bind("<Return>", lambda e: self.enter_callback(e))
        # Shift+Arrow bindings can stay here or move to parent if needed globally
        self.tree.bind("<Shift-Down>", self._on_shift_select_down)
        self.tree.bind("<Shift-Up>", self._on_shift_select_up)
        self.tree.bind("<Control-a>", self._select_all_visible)
        self.tree.bind("<Control-A>", self._select_all_visible)

    def _apply_style(self):
        """Applies ttk styling to the Treeview."""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass  # Use default theme if 'clam' isn't available
        style.configure(
            "Treeview", background=SoP["TREE_FIELD"],
            fieldbackground=SoP["TREE_FIELD"], foreground=SoP["TEXT"], rowheight=28)
        style.configure(
            "Treeview.Heading", background=SoP["ACCENT_DARK"],
            foreground=SoP["ACCENT_HOVER"], relief="flat", font=("Calibri", 11, "bold"))
        style.map("Treeview.Heading", background=[
                  ("active", SoP["ACCENT_DARK"])])
        style.map("Treeview", background=[("selected", SoP["ACCENT"])])

    def populate(self, data_frame: pd.DataFrame):
        """Clears and fills the Treeview with data from a DataFrame."""
        # Clear existing items
        for i in self.tree.get_children():
            self.tree.delete(i)

        # Insert new items
        for _, row in data_frame.iterrows():
            ver = row.get("version", "")
            # Use name as IID, fallback to path if names collide (robustness)
            iid = row["name"]
            if self.tree.exists(iid):
                # Use index as fallback part
                iid = row.get("path", f"{row['name']}_{_}")

            self.tree.insert("", "end", iid=iid, values=(
                row["name"], row["category"], row.get("description", ""), ver))

    # --- Public methods for parent interaction ---
    def get_selection(self) -> Tuple[str, ...]:
        return self.tree.selection()

    def get_focused_item_name(self) -> str | None:
        focus = self.tree.focus()
        if focus:
            return self.tree.item(focus, "values")[0]
        return None

    def get_item_values(self, iid: str) -> Tuple[str, ...]:
        values = self.tree.item(iid, "values")
        if isinstance(values, tuple):
            return tuple(str(v) for v in values)
        else:
            return ()

    def clear_selection(self):
        for sel in self.tree.selection():
            self.tree.selection_remove(sel)

    def focus_first_item(self):
        children = self.tree.get_children()
        if children:
            self.tree.focus(children[0])
            self.tree.selection_set(children[0])
            self.tree.see(children[0])

    # --- Internal Event Handlers ---
    def _on_shift_select_down(self, event):
        focus = self.tree.focus()
        if not focus:
            return "break"
        next_item = self.tree.next(focus)
        if next_item:
            self.tree.selection_add(next_item)
            self.tree.focus(next_item)
            self.tree.see(next_item)
        self.select_callback()  # Notify parent of selection change
        return "break"

    def _on_shift_select_up(self, event):
        focus = self.tree.focus()
        if not focus:
            return "break"
        prev_item = self.tree.prev(focus)
        if prev_item:
            self.tree.selection_add(prev_item)
            self.tree.focus(prev_item)
            self.tree.see(prev_item)
        self.select_callback()  # Notify parent
        return "break"

    def _select_all_visible(self, event):
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children)
        self.select_callback()  # Notify parent
        return "break"
