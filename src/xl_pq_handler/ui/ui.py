# app.py
import os
import sys
import customtkinter as ctk
from tkinter import messagebox

# Import the new manager
from ..classes import PQManager

# Import modularized components
from .theme import SoP
from .views import LibraryView, CreateView, ExtractView


class PQManagerUI:
    def __init__(self, root_path: str, hwnd: int | None = None):
        ctk.set_appearance_mode("Dark")

        self.root = ctk.CTk()
        self.root.title("Shan's PQ Magic ✨")
        self.root.geometry("1200x750")
        self.root.minsize(1000, 700)
        self.root.configure(fg_color=SoP["BG"])

        try:
            icon_path = os.path.join(root_path, "app.ico")
            self.root.iconbitmap(icon_path, icon_path)
        except Exception as e:
            print(f"Could not load icon: {e}")

        # NEW: Initialize the PQManager
        try:
            self.manager = PQManager(root_path, hwnd=hwnd)
        except Exception as e:
            messagebox.showerror(
                "Fatal Error", f"Failed to initialize PQManager:\n{e}")
            self.root.destroy()
            return

        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        self.views = {}

        self._build_layout()
        self.select_view("library")

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _build_layout(self):
        """Builds the main layout: Activity Bar and View Panel."""

        # --- Activity Bar ---
        self.activity_bar = ctk.CTkFrame(
            self.root, width=60, corner_radius=0, fg_color=SoP["FRAME"])
        self.activity_bar.grid(row=0, column=0, sticky="nsw")
        self.activity_bar.grid_rowconfigure(3, weight=1)

        self.nav_btn_library = ctk.CTkButton(
            self.activity_bar, text="📚", font=ctk.CTkFont(size=22), width=50, height=50,
            command=lambda: self.select_view("library"), fg_color="transparent",
            hover_color=SoP["TREE_FIELD"], text_color=SoP["TEXT_DIM"])
        self.nav_btn_library.grid(row=0, column=0, padx=5, pady=(10, 5))

        self.nav_btn_create = ctk.CTkButton(
            self.activity_bar, text="➕", font=ctk.CTkFont(size=22), width=50, height=50,
            command=lambda: self.select_view("create"), fg_color="transparent",
            hover_color=SoP["TREE_FIELD"], text_color=SoP["TEXT_DIM"])
        self.nav_btn_create.grid(row=1, column=0, padx=5, pady=5)

        self.nav_btn_extract = ctk.CTkButton(
            self.activity_bar, text="📥", font=ctk.CTkFont(size=22), width=50, height=50,
            command=lambda: self.select_view("extract"), fg_color="transparent",
            hover_color=SoP["TREE_FIELD"], text_color=SoP["TEXT_DIM"])
        self.nav_btn_extract.grid(row=2, column=0, padx=5, pady=5)

        self.refresh_btn = ctk.CTkButton(
            self.activity_bar, text="🔄", font=ctk.CTkFont(size=22), width=50, height=50,
            command=self.refresh_all_views, fg_color="transparent",
            hover_color=SoP["TREE_FIELD"], text_color=SoP["TEXT_DIM"])
        self.refresh_btn.grid(row=4, column=0, sticky="s", padx=5, pady=10)

        # --- Main Panel (to hold the views) ---
        self.main_panel = ctk.CTkFrame(
            self.root, corner_radius=0, fg_color=SoP["EDITOR"])
        self.main_panel.grid(row=0, column=1, sticky="nsew")
        self.main_panel.grid_columnconfigure(0, weight=1)
        self.main_panel.grid_rowconfigure(0, weight=1)

        # --- Instantiate Views ---
        # Pass callbacks to the views so they can trigger actions in the main app
        self.views["library"] = LibraryView(
            self.main_panel, self.manager, self.refresh_all_views)
        self.views["create"] = CreateView(
            self.main_panel,
            self.manager,
            refresh_callback=self.refresh_all_views,
            switch_to_library_callback=lambda: self.select_view("library")
        )
        self.views["extract"] = ExtractView(
            self.main_panel,
            self.manager,
            refresh_callback=self.refresh_all_views
        )

        # Place views in the grid, they will be hidden/shown by select_view
        for view in self.views.values():
            view.grid(row=0, column=0, sticky="nsew", padx=20, pady=15)

    def select_view(self, view_name: str):
        """Shows the selected view and hides the others."""

        # Reset all button colors
        self.nav_btn_library.configure(
            fg_color="transparent", text_color=SoP["TEXT_DIM"])
        self.nav_btn_create.configure(
            fg_color="transparent", text_color=SoP["TEXT_DIM"])
        self.nav_btn_extract.configure(
            fg_color="transparent", text_color=SoP["TEXT_DIM"])

        # Hide all views
        for view in self.views.values():
            view.grid_remove()

        # Show the selected view and highlight the button
        if view_name == "library":
            self.views["library"].grid()
            self.nav_btn_library.configure(
                fg_color=SoP["TREE_FIELD"], text_color=SoP["ACCENT"])
        elif view_name == "create":
            self.views["create"].grid()
            self.nav_btn_create.configure(
                fg_color=SoP["TREE_FIELD"], text_color=SoP["ACCENT"])
        elif view_name == "extract":
            self.views["extract"].grid()
            self.nav_btn_extract.configure(
                fg_color=SoP["TREE_FIELD"], text_color=SoP["ACCENT"])

    def refresh_all_views(self):
        """
        Callback function to rebuild the index and refresh all views
        that depend on it.
        """
        try:
            self.refresh_btn.configure(text_color=SoP["ACCENT"])
            self.manager.build_index()
            # Tell LibraryView to reload its data
            self.views["library"].refresh_data()
            # Tell CreateView to clear its form
            self.views["create"].clear_form()

            messagebox.showinfo("Refresh Complete",
                                "The query index has been rebuilt.")
            self.refresh_btn.configure(text_color=SoP["TEXT_DIM"])
        except Exception as e:
            messagebox.showerror(
                "Refresh Error", f"Failed to rebuild index or refresh UI:\n{e}")
            self.refresh_btn.configure(text_color=SoP["TEXT_DIM"])

    def _on_close(self):
        """Handle window close event."""
        self.root.destroy()
