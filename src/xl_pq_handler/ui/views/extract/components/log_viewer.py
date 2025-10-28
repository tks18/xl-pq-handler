import customtkinter as ctk

from ....theme import SoP


class LogViewer(ctk.CTkFrame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.log_textbox = ctk.CTkTextbox(
            self, border_color=SoP["FRAME"], border_width=2,
            fg_color=SoP["FRAME"], text_color=SoP["TEXT_DIM"]
        )
        self.log_textbox.insert("1.0", "Extraction log will appear here...")
        self.log_textbox.configure(state="disabled")
        self.log_textbox.grid(row=0, column=0, sticky="nsew")

        # Configure log tags
        self.log_textbox.tag_config("accent", foreground=SoP["ACCENT"])
        self.log_textbox.tag_config(
            "accent_bold", foreground=SoP["ACCENT_HOVER"])
        self.log_textbox.tag_config("error", foreground="#FF5555")

    def clear_log(self):
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

    def append_log(self, message, tag=None):
        # No need for after() here, parent view will call this from main thread
        self.log_textbox.configure(state="normal")
        if tag:
            self.log_textbox.insert("end", message, (tag,))
        else:
            self.log_textbox.insert("end", message)
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")
