# ui_create_view.py
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import threading
import os
from typing import Callable

from ...classes.pq_manager import PQManager
from ...classes.pq_manager.models import PowerQueryScript, PowerQueryMetadata
from ..theme import SoP


class CreateView(ctk.CTkFrame):
    """
    Manages the 'Create' view for making new .pq files.
    """

    def __init__(self, parent, manager: PQManager,
                 refresh_callback: Callable,
                 switch_to_library_callback: Callable):
        super().__init__(parent, fg_color="transparent")
        self.manager = manager
        self.refresh_callback = refresh_callback
        self.switch_to_library_callback = switch_to_library_callback

        self._build_widgets()

    def _build_widgets(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        title = ctk.CTkLabel(
            self, text="Create New Power Query",
            font=ctk.CTkFont(size=24, weight="bold"), text_color=SoP["ACCENT_HOVER"])
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))

        form_frame = ctk.CTkScrollableFrame(
            self, fg_color=SoP["FRAME"], corner_radius=8)
        form_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        form_frame.grid_columnconfigure(1, weight=1)

        def create_form_row(parent, label, row):
            ctk.CTkLabel(parent, text=label, text_color=SoP["TEXT_DIM"]).grid(
                row=row, column=0, sticky="w", padx=10, pady=8)
            entry = ctk.CTkEntry(
                parent, border_color=SoP["TREE_FIELD"],
                fg_color=SoP["EDITOR"], text_color=SoP["TEXT"])
            entry.grid(row=row, column=1, sticky="ew", padx=10, pady=8)
            return entry

        self.create_entry_name = create_form_row(form_frame, "Name*", 0)
        self.create_entry_category = create_form_row(form_frame, "Category", 1)
        self.create_entry_version = create_form_row(form_frame, "Version", 2)
        self.create_entry_tags = create_form_row(form_frame, "Tags (csv)", 3)
        self.create_entry_deps = create_form_row(
            form_frame, "Dependencies (csv)", 4)

        ctk.CTkLabel(form_frame, text="Description", text_color=SoP["TEXT_DIM"]).grid(
            row=5, column=0, sticky="w", padx=10, pady=8)
        self.create_text_desc = ctk.CTkTextbox(
            form_frame, height=80, border_color=SoP["TREE_FIELD"],
            fg_color=SoP["EDITOR"], text_color=SoP["TEXT"])
        self.create_text_desc.grid(
            row=5, column=1, sticky="ew", padx=10, pady=8)

        ctk.CTkLabel(form_frame, text="Query Body*", text_color=SoP["TEXT_DIM"]).grid(
            row=6, column=0, sticky="nw", padx=10, pady=8)
        self.create_text_body = ctk.CTkTextbox(
            form_frame, height=250, border_color=SoP["TREE_FIELD"],
            fg_color=SoP["EDITOR"], text_color=SoP["TEXT"], font=("Consolas", 12))
        self.create_text_body.grid(
            row=6, column=1, sticky="ew", padx=10, pady=8)

        self.create_save_btn = ctk.CTkButton(
            self, text="💾 Save New Query", height=40,
            command=self._threaded_create_new_pq, fg_color=SoP["ACCENT"],
            hover_color=SoP["ACCENT_HOVER"], text_color="#000000",
            font=ctk.CTkFont(weight="bold"))
        self.create_save_btn.grid(
            row=2, column=0, columnspan=2, sticky="e", pady=20, padx=0)

    def clear_form(self):
        """Called by the main app after a successful refresh."""
        self.create_entry_name.delete(0, "end")
        self.create_entry_category.delete(0, "end")
        self.create_entry_version.delete(0, "end")
        self.create_entry_tags.delete(0, "end")
        self.create_entry_deps.delete(0, "end")
        self.create_text_desc.delete("1.0", "end")
        self.create_text_body.delete("1.0", "end")

    def _threaded_create_new_pq(self):
        name = self.create_entry_name.get().strip()
        body = self.create_text_body.get("1.0", "end").strip()
        if not name or not body:
            messagebox.showerror("Error", "Name and Query Body are required.")
            return

        category = self.create_entry_category.get().strip() or "Uncategorized"
        version = self.create_entry_version.get().strip() or "1.0"
        desc = self.create_text_desc.get("1.0", "end").strip()

        def split_csv(csv_str):
            return [tag.strip() for tag in csv_str.split(",") if tag.strip()]
        tags = split_csv(self.create_entry_tags.get())
        deps = split_csv(self.create_entry_deps.get())

        threading.Thread(
            target=self._create_new_pq,
            args=(name, body, category, desc, tags, deps, version),
            daemon=True
        ).start()

    def _create_new_pq(self, name, body, category, desc, tags, deps, version):
        """
        Thread target to create a new PQ file using the new manager.
        """
        try:
            # NEW API LOGIC: Construct the script object and save it.

            # 1. Determine safe path (logic from old handler)
            safe_name = "".join(c for c in name if c.isalnum()
                                or c in (' ', '_', '-')).rstrip()
            safe_category = "".join(
                c for c in category if c.isalnum() or c in (' ', '_', '-')).rstrip()
            if not safe_category:
                safe_category = "Uncategorized"

            # Use the manager's root path
            target_dir = os.path.join(self.manager.store.root, safe_category)
            out_path = os.path.join(target_dir, f"{safe_name}.pq")

            # 2. Create Pydantic Models
            meta = PowerQueryMetadata(
                name=name,
                category=category,
                description=desc,
                tags=tags,
                dependencies=deps,
                version=version,
                path=out_path
            )

            script = PowerQueryScript(meta=meta, body=body)

            # 3. Save using the store
            self.manager.store.save_script(script, overwrite=True)

            # 4. Rebuild the index *after* saving
            self.manager.build_index()

            # 5. Use callbacks to update UI on the main thread
            self.master.after(0, lambda: messagebox.showinfo(
                "Success", f"Query '{name}' was created successfully."))
            self.master.after(0, self.refresh_callback)
            self.master.after(0, self.switch_to_library_callback)

        except Exception as e:
            self.master.after(0, lambda: messagebox.showerror(
                "Error", f"Failed to create query:\n{e}"))
