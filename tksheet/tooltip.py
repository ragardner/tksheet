from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Any, Literal

from .constants import ctrl_key


class TooltipTkText(tk.Text):
    """Custom Text widget for the Tooltip class."""

    def __init__(
        self,
        parent: tk.Misc,
        rc_bindings: list[str] = "<3>",
    ) -> None:
        super().__init__(
            parent,
            spacing1=0,
            spacing2=1,
            spacing3=0,
            bd=0,
            highlightthickness=0,
            undo=True,
            maxundo=30,
        )
        self.parent = parent
        self.rc_popup_menu = tk.Menu(self, tearoff=0)
        self.bind("<1>", lambda event: self.focus_set())
        for b in rc_bindings:
            self.bind(b, self.rc)
        self.bind(f"<{ctrl_key}-a>", self.select_all)
        self.bind(f"<{ctrl_key}-A>", self.select_all)
        self.bind("<Delete>", self.delete_key)

    def reset(
        self,
        menu_kwargs: dict,
        sheet_ops: dict,
        bg: str,
        fg: str,
        select_bg: str,
        select_fg: str,
        text: str,
        readonly: bool,
    ) -> None:
        """Reset the text widget's appearance and menu options."""
        self.config(
            font=sheet_ops.table_font,
            background=bg,
            foreground=fg,
            insertbackground=fg,
            selectbackground=select_bg,
            selectforeground=select_fg,
        )
        self.editor_del_key = sheet_ops.editor_del_key
        self.config(state="normal")
        self.delete(1.0, "end")
        self.insert(1.0, text)
        self.edit_reset()
        self.rc_popup_menu.delete(0, "end")
        self.rc_popup_menu.add_command(
            label=sheet_ops.select_all_label,
            accelerator=sheet_ops.select_all_accelerator,
            image=sheet_ops.select_all_image,
            compound=sheet_ops.select_all_compound,
            command=self.select_all,
            **menu_kwargs,
        )
        if not readonly:
            self.rc_popup_menu.add_command(
                label=sheet_ops.cut_label,
                accelerator=sheet_ops.cut_accelerator,
                image=sheet_ops.cut_image,
                compound=sheet_ops.cut_compound,
                command=self.cut,
                **menu_kwargs,
            )
        self.rc_popup_menu.add_command(
            label=sheet_ops.copy_label,
            accelerator=sheet_ops.copy_accelerator,
            image=sheet_ops.copy_image,
            compound=sheet_ops.copy_compound,
            command=self.copy,
            **menu_kwargs,
        )
        if not readonly:
            self.rc_popup_menu.add_command(
                label=sheet_ops.paste_label,
                accelerator=sheet_ops.paste_accelerator,
                image=sheet_ops.paste_image,
                compound=sheet_ops.paste_compound,
                command=self.paste,
                **menu_kwargs,
            )
            self.rc_popup_menu.add_command(
                label=sheet_ops.undo_label,
                accelerator=sheet_ops.undo_accelerator,
                image=sheet_ops.undo_image,
                compound=sheet_ops.undo_compound,
                command=self.undo,
                **menu_kwargs,
            )
            self.rc_popup_menu.add_command(
                label=sheet_ops.redo_label,
                accelerator=sheet_ops.redo_accelerator,
                image=sheet_ops.redo_image,
                compound=sheet_ops.redo_compound,
                command=self.redo,
                **menu_kwargs,
            )

    def rc(self, event: Any) -> None:
        """Show the right-click popup menu."""
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)
        self.focus_set()

    def delete_key(self, event: Any = None) -> None:
        """Handle the Delete key based on editor configuration."""
        if self.editor_del_key == "forward":
            return
        elif not self.editor_del_key:
            return "break"
        elif self.editor_del_key == "backward":
            if self.tag_ranges("sel"):
                return
            if self.index("insert") == "1.0":
                return "break"
            self.delete("insert-1c")
            return "break"

    def select_all(self, event: Any = None) -> Literal["break"]:
        """Select all text in the widget."""
        self.tag_add(tk.SEL, "1.0", "end-1c")
        self.mark_set(tk.INSERT, "end-1c")
        return "break"

    def cut(self, event: Any = None) -> Literal["break"]:
        """Cut selected text."""
        self.event_generate(f"<{ctrl_key}-x>")
        self.event_generate("<KeyRelease>")
        return "break"

    def copy(self, event: Any = None) -> Literal["break"]:
        """Copy selected text."""
        self.event_generate(f"<{ctrl_key}-c>")
        return "break"

    def paste(self, event: Any = None) -> Literal["break"]:
        """Paste text from clipboard."""
        self.event_generate(f"<{ctrl_key}-v>")
        self.event_generate("<KeyRelease>")
        return "break"

    def undo(self, event: Any = None) -> Literal["break"]:
        """Undo the last action."""
        self.event_generate(f"<{ctrl_key}-z>")
        self.event_generate("<KeyRelease>")
        return "break"

    def redo(self, event: Any = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-Shift-z>")
        self.event_generate("<KeyRelease>")
        return "break"


class Tooltip(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        sheet_ops: dict,
        menu_kwargs: dict,
        bg: str,
        fg: str,
        select_bg: str,
        select_fg: str,
        scrollbar_style: str,
        rc_bindings: list[str] = "<3>",
    ) -> None:
        super().__init__(parent)
        self.withdraw()  # Hide until positioned
        self.overrideredirect(True)  # Borderless window
        self.cell_readonly = True
        self.note_readonly = True

        # Store parameters
        self.sheet_ops = sheet_ops
        self.menu_kwargs = menu_kwargs
        self.font = sheet_ops.table_font
        self.bg = bg
        self.fg = fg
        self.select_bg = select_bg
        self.select_fg = select_fg
        self.row = 0
        self.col = 0

        # Create border frame for visual distinction
        self.border_frame = tk.Frame(self, background=bg)
        self.border_frame.pack(fill="both", expand=True, padx=1, pady=1)

        # Create notebook, but don’t pack it yet
        self.notebook = ttk.Notebook(self.border_frame)

        # Content frame as child of border_frame
        self.content_frame = ttk.Frame(self.border_frame)
        self.content_text = TooltipTkText(self.content_frame, rc_bindings=rc_bindings)
        self.content_scrollbar = ttk.Scrollbar(self.content_frame, orient="vertical", style=scrollbar_style)
        self.content_scrollbar.pack(side="right", fill="y")
        self.content_text.pack(side="left", fill="both", expand=True)
        self.content_scrollbar.configure(command=self.content_text.yview)
        self.content_text.configure(yscrollcommand=self.content_scrollbar.set)

        # Note frame as child of border_frame
        self.note_frame = ttk.Frame(self.border_frame)
        self.note_text = TooltipTkText(self.note_frame, rc_bindings=rc_bindings)
        self.note_scrollbar = ttk.Scrollbar(self.note_frame, orient="vertical", style=scrollbar_style)
        self.note_scrollbar.pack(side="right", fill="y")
        self.note_text.pack(side="left", fill="both", expand=True)
        self.note_scrollbar.configure(command=self.note_text.yview)
        self.note_text.configure(yscrollcommand=self.note_scrollbar.set)

    def setup_note_only_mode(self) -> None:
        """Configure the tooltip to show only the note text widget."""
        self.notebook.pack_forget()  # Remove notebook from layout
        for tab in self.notebook.tabs():  # Clear all tabs to free the frames
            self.notebook.forget(tab)
        self.content_frame.pack_forget()  # Ensure content_frame is not directly packed
        self.note_frame.pack(fill="both", expand=True)  # Show note_frame directly

    def setup_single_text_mode(self) -> None:
        """Configure the tooltip to show only the content text widget."""
        self.notebook.pack_forget()  # Remove notebook from layout
        for tab in self.notebook.tabs():  # Clear all tabs
            self.notebook.forget(tab)
        self.note_frame.pack_forget()  # Ensure note_frame is not directly packed
        self.content_frame.pack(fill="both", expand=True)  # Show content_frame directly

    def setup_notebook_mode(self) -> None:
        """Configure the tooltip to show a notebook with Cell and Note tabs."""
        self.content_frame.pack_forget()  # Ensure content_frame is not directly packed
        self.note_frame.pack_forget()  # Ensure note_frame is not directly packed
        self.notebook.add(self.content_frame, text="Cell")  # Add content tab
        self.notebook.add(self.note_frame, text="Note")  # Add note tab
        self.notebook.pack(fill="both", expand=True)  # Show notebook
        self.notebook.select(0)  # Select the "Cell" tab by default

    def reset(
        self,
        text: str,
        cell_readonly: bool,
        note: str | None,
        note_readonly: bool,
        row: int,
        col: int,
        menu_kwargs: dict,
        bg: str,
        fg: str,
        select_bg: str,
        select_fg: str,
        user_can_create_notes: bool,
        note_only: bool,
        width: int,
        height: int,
    ) -> None:
        self.cell_readonly = cell_readonly
        self.note_readonly = note_readonly
        self.menu_kwargs = menu_kwargs
        self.font = self.sheet_ops.table_font
        self.bg = bg
        self.fg = fg
        self.select_bg = select_bg
        self.select_fg = select_fg
        self.row = row
        self.col = col
        self.config(bg=self.sheet_ops.table_selected_box_cells_fg)
        self.border_frame.config(background=bg)
        self.content_text.configure(
            wrap="word",
            width=30,
            height=5,
            state="disabled" if cell_readonly else "normal",
        )
        kws = {
            "menu_kwargs": menu_kwargs,
            "sheet_ops": self.sheet_ops,
            "bg": bg,
            "fg": fg,
            "select_bg": select_bg,
            "select_fg": select_fg,
        }
        self.content_text.reset(**kws, text=text, readonly=cell_readonly)
        self.content_text.configure(state="disabled" if cell_readonly else "normal")
        self.note_text.configure(
            wrap="word",
            width=30,
            height=5,
        )
        self.note_text.reset(**kws, text="" if note is None else note, readonly=note_readonly)
        self.note_text.configure(state="disabled" if note_readonly else "normal")
        # Set up UI based on condition
        if note_only:
            self.setup_note_only_mode()
        elif note is None and not user_can_create_notes:
            self.setup_single_text_mode()
        else:
            self.setup_notebook_mode()
        self.adjust_size(width, height)

    def adjust_size(self, width, height) -> None:
        """Adjust tooltip size to given dimensions."""
        self.update_idletasks()
        self.geometry(f"{width}x{height}")

    def set_position(self, x_root: int, y_root: int) -> None:
        """Position tooltip near the cell, avoiding screen edges."""
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = self.winfo_width()
        height = self.winfo_height()

        # Adjust position to avoid screen edges
        x = min(x_root, screen_width - width)
        y = min(y_root, screen_height - height)
        self.geometry(f"+{x}+{y}")
        self.deiconify()

    def get(self) -> tuple[int, int, str, str]:
        return (
            self.row,
            self.col,
            self.content_text.get("1.0", "end-1c"),
            self.note_text.get("1.0", "end-1c"),
        )

    def destroy(self) -> None:
        super().destroy()
