import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import os
from typing import Callable

from .....classes import PQManager, PowerQueryMetadata, PowerQueryScript
from ....theme import SoP
from ....components import CTkCodeView


class EditMetadataDialog(ctk.CTkToplevel):
    """
    A Toplevel window dialog for editing the metadata of a Power Query script.
    """

    def __init__(self, parent, manager: PQManager, script_to_edit: PowerQueryScript, refresh_callback: Callable):
        super().__init__(parent)
        self.manager = manager
        self.original_script = script_to_edit
        # Callback to refresh the main library view
        self.refresh_callback = refresh_callback

        self.title(f"Edit: {self.original_script.meta.name}")
        self.geometry("700x700")
        self.transient(parent)  # Keep on top of parent
        self.grab_set()      # Block interaction with parent
        self.configure(fg_color=SoP["BG"])

        self._build_widgets()
        self._prefill_form()

    def _build_widgets(self):
        """Builds the form elements for the dialog."""
        form_frame = ctk.CTkScrollableFrame(
            self, fg_color=SoP["FRAME"], corner_radius=8)
        form_frame.pack(fill="both", expand=True, padx=15, pady=15)
        form_frame.grid_columnconfigure(1, weight=1)

        def create_form_row(parent, label, row):
            ctk.CTkLabel(parent, text=label, text_color=SoP["TEXT_DIM"], font=ctk.CTkFont(family="Segoe UI")).grid(
                row=row, column=0, sticky="w", padx=10, pady=8)
            entry = ctk.CTkEntry(
                parent, border_color=SoP["TREE_FIELD"],
                fg_color=SoP["EDITOR"], text_color=SoP["TEXT"], font=ctk.CTkFont(family="Segoe UI"))
            entry.grid(row=row, column=1, sticky="ew", padx=10, pady=8)
            return entry

        self.entry_name = create_form_row(form_frame, "Name*", 0)
        self.entry_category = create_form_row(form_frame, "Category", 1)
        self.entry_version = create_form_row(form_frame, "Version", 2)
        self.entry_tags = create_form_row(form_frame, "Tags (csv)", 3)

        # Dependencies Row (with Auto-Detect)
        ctk.CTkLabel(form_frame, text="Dependencies (csv)", text_color=SoP["TEXT_DIM"], font=ctk.CTkFont(family="Segoe UI")).grid(
            row=4, column=0, sticky="w", padx=10, pady=8)
        dep_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        dep_frame.grid(row=4, column=1, sticky="ew")
        dep_frame.grid_columnconfigure(0, weight=1)
        self.entry_deps = ctk.CTkEntry(
            dep_frame, border_color=SoP["TREE_FIELD"],
            fg_color=SoP["EDITOR"], text_color=SoP["TEXT"], font=ctk.CTkFont(family="Segoe UI"))
        self.entry_deps.grid(row=0, column=0, sticky="ew", padx=10, pady=8)
        auto_detect_btn = ctk.CTkButton(
            dep_frame, text="Auto-Detect", width=100,
            fg_color=SoP["TREE_FIELD"], hover_color=SoP["ACCENT_DARK"],
            command=self._auto_detect_deps, font=ctk.CTkFont(family="Segoe UI")
        )
        auto_detect_btn.grid(row=0, column=1, sticky="e")

        # Description Textbox
        ctk.CTkLabel(form_frame, text="Description", text_color=SoP["TEXT_DIM"], font=ctk.CTkFont(family="Segoe UI")).grid(
            row=5, column=0, sticky="w", padx=10, pady=8)
        self.text_desc = ctk.CTkTextbox(
            form_frame, height=80, border_color=SoP["TREE_FIELD"],
            fg_color=SoP["EDITOR"], text_color=SoP["TEXT"], font=ctk.CTkFont(family="Segoe UI"))
        self.text_desc.grid(row=5, column=1, sticky="ew", padx=10, pady=8)

        # Read-Only Code View
        ctk.CTkLabel(form_frame, text="Query Body (Read-Only)", text_color=SoP["TEXT_DIM"], font=ctk.CTkFont(family="Segoe UI")).grid(
            row=6, column=0, sticky="nw", padx=10, pady=8)
        self.code_view = CTkCodeView(
            form_frame,
            manager=self.manager,
            height=200
        )
        self.code_view.grid(row=6, column=1, sticky="nsew", padx=10, pady=8)

        # Save Button
        save_btn = ctk.CTkButton(
            self, text="💾 Save Changes", height=40,
            command=self._on_save, fg_color="transparent", border_width=1, border_color=SoP["ACCENT"],
            hover_color=SoP["TREE_FIELD"], text_color=SoP["ACCENT"], font=ctk.CTkFont(family="Segoe UI"))
        save_btn.pack(side="bottom", fill="x", padx=15, pady=15)

    def _prefill_form(self):
        """Fills the form widgets with the original script data."""
        script = self.original_script
        self.entry_name.insert(0, script.meta.name)
        self.entry_category.insert(0, script.meta.category)
        self.entry_version.insert(0, script.meta.version)
        self.entry_tags.insert(0, ", ".join(script.meta.tags))
        self.entry_deps.insert(0, ", ".join(script.meta.dependencies))
        self.text_desc.insert("1.0", script.meta.description)
        self.code_view.set_code(script.body)  # Use the codeview's method

    def _auto_detect_deps(self):
        """Auto-detects dependencies from the code view."""
        try:
            # Get body directly from the code_view's internal textbox
            body = self.code_view.get("1.0", "end")
            deps_list = self.manager.get_dependencies_from_code(body)

            if not deps_list:
                messagebox.showinfo(
                    "Auto-Detect", "No dependencies found.", parent=self)
                return

            self.entry_deps.delete(0, "end")
            self.entry_deps.insert(0, ", ".join(deps_list))

        except Exception as e:
            messagebox.showerror(
                "Parse Error", f"Failed to parse code:\n{e}", parent=self)

    def _on_save(self):
        """Handles the save action."""
        try:
            # 1. Get new values
            new_name = self.entry_name.get().strip()
            if not new_name:
                raise ValueError("Name cannot be empty.")
            new_category = self.entry_category.get().strip() or "Uncategorized"

            # 2. Construct the new path
            safe_name = "".join(c for c in new_name if c.isalnum()
                                or c in (' ', '_', '-')).rstrip()
            safe_category = "".join(c for c in new_category if c.isalnum() or c in (
                ' ', '_', '-')).rstrip() or "Uncategorized"
            new_path = os.path.join(
                self.manager.store.root, safe_category, f"{safe_name}.pq")

            # 3. Create the new PowerQueryMetadata object
            new_meta = PowerQueryMetadata(
                name=new_name,
                category=new_category,
                version=self.entry_version.get().strip() or "1.0",
                tags=[t.strip()
                      for t in self.entry_tags.get().split(",") if t.strip()],
                dependencies=[d.strip()
                              for d in self.entry_deps.get().split(",") if d.strip()],
                description=self.text_desc.get("1.0", "end").strip(),
                path=os.path.abspath(new_path)
            )

            # 4. Call the manager to do the update
            self.manager.update_query_metadata(
                self.original_script.meta.name, new_meta)

            # 5. Call the refresh callback provided by the parent view
            self.refresh_callback()

            self.destroy()  # Close the dialog

        except Exception as e:
            messagebox.showerror(
                "Save Error", f"Failed to save changes:\n{e}", parent=self)
