from __future__ import annotations

import re
import tkinter as tk
from collections.abc import Callable
from typing import Any, Literal

from .constants import alt_key, ctrl_key
from .functions import recursive_bind
from .other_classes import DotDict, FontTuple


class FindWindowTkText(tk.Text):
    """Custom Text widget for the FindWindow class."""

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
        font: FontTuple,
        bg: str,
        fg: str,
        select_bg: str,
        select_fg: str,
    ) -> None:
        """Reset the text widget's appearance and menu options."""
        self.config(
            font=font,
            background=bg,
            foreground=fg,
            insertbackground=fg,
            selectbackground=select_bg,
            selectforeground=select_fg,
        )
        self.editor_del_key = sheet_ops.editor_del_key
        self.rc_popup_menu.delete(0, "end")
        self.rc_popup_menu.add_command(
            label=sheet_ops.select_all_label,
            accelerator=sheet_ops.select_all_accelerator,
            image=sheet_ops.select_all_image,
            compound=sheet_ops.select_all_compound,
            command=self.select_all,
            **menu_kwargs,
        )
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
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

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


class LabelTooltip(tk.Toplevel):
    def __init__(self, parent: tk.Misc, text: str, bg: str, fg: str) -> None:
        super().__init__(parent)
        self.withdraw()
        self.overrideredirect(True)
        self.label = tk.Label(self, text=text, background=bg, foreground=fg, relief="flat", borderwidth=0)
        self.label.pack()
        self.text = text
        self.config(background=bg, highlightbackground=bg, highlightthickness=0)
        self.update_idletasks()


class FindWindow(tk.Frame):
    """A frame containing find and replace functionality with label highlighting and tooltips."""

    def __init__(
        self,
        parent: tk.Misc,
        find_next_func: Callable,
        find_prev_func: Callable,
        close_func: Callable,
        replace_func: Callable,
        replace_all_func: Callable,
        toggle_replace_func: Callable,
        drag_func: Callable,
        rc_bindings: list[str] = "<3>",
    ) -> None:
        super().__init__(
            parent,
            width=0,
            height=0,
            bd=0,
        )
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(4, uniform="group1")
        self.grid_columnconfigure(5, uniform="group2")
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self.grid_propagate(False)
        self.parent = parent
        self.tooltip_after_id = None
        self.tooltip_last_x = None
        self.tooltip_last_y = None
        self.tooltip_widget = None  # Added to track the current widget
        self.tooltip = None

        self.find_next_func = find_next_func
        self.find_prev_func = find_prev_func
        self.replace_func = replace_func
        self.toggle_replace_func = toggle_replace_func
        self.drag_func = drag_func
        self.close_func = close_func

        self.toggle_replace = tk.Label(self, text="↓", cursor="sb_h_double_arrow", highlightthickness=1)
        self.toggle_replace.grid(row=0, column=0, sticky="ns")
        self.toggle_replace.grid_remove()

        self.tktext = FindWindowTkText(self, rc_bindings=rc_bindings)
        self.tktext.grid(row=0, column=1, sticky="nswe")

        self.find_previous_arrow = tk.Label(self, text="↑", cursor="hand2", highlightthickness=1)
        self.find_previous_arrow.grid(row=0, column=2)

        self.find_next_arrow = tk.Label(self, text="↓", cursor="hand2", highlightthickness=1)
        self.find_next_arrow.grid(row=0, column=3)

        self.find_in_selection = False
        self.in_selection = tk.Label(self, text="≡", cursor="hand2", highlightthickness=1)
        self.in_selection.grid(row=0, column=4)

        self.close = tk.Label(self, text="✕", cursor="hand2", highlightthickness=1)
        self.close.grid(row=0, column=5, sticky="nswe")

        self.separator = tk.Frame(self, height=1)
        self.separator.grid(row=1, column=1, columnspan=3, sticky="we")
        self.separator.grid_remove()

        self.replace_tktext = FindWindowTkText(self, rc_bindings=rc_bindings)
        self.replace_tktext.grid(row=2, column=1, columnspan=4, sticky="nswe")
        self.replace_tktext.grid_remove()

        self.replace_next = tk.Label(self, text="→", cursor="hand2", highlightthickness=1)
        self.replace_next.grid(row=2, column=4, sticky="nswe")
        self.replace_next.grid_remove()

        self.replace_all = tk.Label(self, text="⟳", cursor="hand2", highlightthickness=1)
        self.replace_all.grid(row=2, column=5, sticky="nswe")
        self.replace_all.grid_remove()

        self.tktext.bind("<Tab>", self.handle_tab)
        self.replace_tktext.bind("<Tab>", self.handle_tab)
        self.tktext.bind("<Return>", self.handle_return)
        self.replace_tktext.bind("<Return>", self.handle_return)

        self.bind_label(self.toggle_replace, self.toggle_replace_window, self.drag_func)
        self.bind_label(self.find_previous_arrow, find_prev_func)
        self.bind_label(self.find_next_arrow, find_next_func)
        self.bind_label(self.in_selection, self.toggle_in_selection)
        self.bind_label(self.close, close_func)
        self.bind_label(self.replace_next, replace_func)
        self.bind_label(self.replace_all, replace_all_func)

        self.replace_visible = False
        self.bg = None
        self.fg = None
        self.pressed_label = None

        for c in ("l", "L"):
            recursive_bind(self, f"<{alt_key}-{c}>", self.toggle_in_selection)

        action_labels = [
            (self.toggle_replace, "Toggle Replace"),
            (self.find_previous_arrow, "Previous Match"),
            (self.find_next_arrow, "Next Match"),
            (self.in_selection, "Find in Selection"),
            (self.close, "Close"),
            (self.replace_next, "Replace"),
            (self.replace_all, "Replace All"),
        ]
        for widget, text in action_labels:
            widget.tooltip_text = text
            widget.bind("<Enter>", self.on_enter)
            widget.bind("<Leave>", self.on_leave)

    def bind_label(self, label: tk.Label, func: Callable, motion_func: Callable | None = None) -> None:
        """Bind press, release, and optional motion events with highlight changes."""

        def on_press(event: tk.Event) -> None:
            label.config(highlightbackground=self.border_color, highlightcolor=self.border_color)
            self.pressed_label = label

        def on_release(event: tk.Event) -> None:
            self.pressed_label = None
            if 0 <= event.x < label.winfo_width() and 0 <= event.y < label.winfo_height():
                label.config(highlightbackground=self.fg, highlightcolor=self.fg)
                func(event)
            else:
                label.config(highlightbackground=self.bg, highlightcolor=self.fg)

        label.bind("<Button-1>", on_press)
        label.bind("<ButtonRelease-1>", on_release)
        if motion_func:
            label.bind("<B1-Motion>", motion_func)

    def on_enter(self, event: tk.Event) -> None:
        """Handle mouse entering a widget."""
        widget = event.widget
        self.enter_label(widget)
        self.tooltip_widget = widget
        self.tooltip_last_x, self.tooltip_last_y = get_mouse_coords(widget)
        self.start_tooltip_timer()

    def on_leave(self, event: tk.Event) -> None:
        """Handle mouse leaving a widget."""
        widget = event.widget
        self.leave_label(widget)
        self.hide_tooltip()
        self.cancel_tooltip()
        self.tooltip_widget = None

    def enter_label(self, widget: tk.Misc) -> None:
        """Highlight label on hover if no label is pressed."""
        if self.pressed_label is None:
            widget.config(highlightbackground=self.fg, highlightcolor=self.fg)

    def leave_label(self, widget: tk.Misc) -> None:
        """Remove highlight on leave unless toggled or pressed."""
        if self.pressed_label is None:
            if widget == self.in_selection and self.find_in_selection:
                return
            widget.config(highlightbackground=self.bg, highlightcolor=self.fg)

    def focus_find(self, event: tk.Misc = None) -> Literal["break"]:
        widget = self.focus_get()
        if widget == self.tktext:
            self.tktext.select_all()
        else:
            self.tktext.focus_set()
        return "break"

    def focus_replace(self, event: tk.Misc = None) -> Literal["break"]:
        if self.replace_enabled and not self.replace_visible:
            self.toggle_replace_window()
        widget = self.focus_get()
        if widget == self.replace_tktext:
            self.replace_tktext.select_all()
        else:
            self.replace_tktext.focus_set()
        return "break"

    def toggle_replace_window(self, event: tk.Misc = None) -> None:
        """Toggle visibility of the replace window."""
        if self.replace_visible:
            self.replace_tktext.grid_remove()
            self.replace_next.grid_remove()
            self.replace_all.grid_remove()
            self.separator.grid_remove()
            self.toggle_replace.config(text="↓")
            self.toggle_replace.grid(row=0, column=0, rowspan=1, sticky="ns")
            self.replace_visible = False
        elif self.replace_enabled:
            self.separator.grid()
            self.replace_tktext.grid()
            self.replace_next.grid()
            self.replace_all.grid()
            self.toggle_replace.config(text="↑")
            self.toggle_replace.grid(row=0, column=0, rowspan=3, sticky="ns")
            self.replace_visible = True
        self.toggle_replace_func()

    def toggle_in_selection(self, event: tk.Misc) -> None:
        """Toggle the find-in-selection state."""
        self.find_in_selection = not self.find_in_selection
        self.enter_label(self.in_selection)
        self.leave_label(self.in_selection)

    def handle_tab(self, event: tk.Event) -> Literal["break"]:
        """Switch focus between find and replace text widgets."""
        if not self.replace_visible:
            self.toggle_replace_window()
        if event.widget == self.tktext:
            self.replace_tktext.focus_set()
        elif event.widget == self.replace_tktext:
            self.tktext.focus_set()
        return "break"

    def handle_return(self, event: tk.Event) -> Literal["break"]:
        """Trigger find or replace based on focused widget."""
        if event.widget == self.tktext:
            self.find_next_func()
        elif event.widget == self.replace_tktext:
            self.replace_func()
        return "break"

    def get(self) -> str:
        """Return the find text."""
        return self.tktext.get("1.0", "end-1c")

    def get_replace(self) -> str:
        """Return the replace text."""
        return self.replace_tktext.get("1.0", "end-1c")

    def get_num_lines(self) -> int:
        """Return the number of lines in the find text."""
        return int(self.tktext.index("end-1c").split(".")[0])

    def set_text(self, text: str = "") -> None:
        """Set the find text."""
        self.tktext.delete(1.0, "end")
        self.tktext.insert(1.0, text)

    def reset(
        self,
        border_color: str,
        grid_color: str,
        menu_kwargs: DotDict,
        sheet_ops: DotDict,
        bg: str,
        fg: str,
        select_bg: str,
        select_fg: str,
        replace_enabled: bool,
    ) -> None:
        """Reset styles and configurations."""
        self.replace_enabled = replace_enabled
        if replace_enabled:
            self.toggle_replace.grid()
        else:
            self.toggle_replace.grid_remove()
        self.bg = bg
        self.fg = fg
        self.border_color = border_color
        self.tktext.reset(
            menu_kwargs=menu_kwargs,
            sheet_ops=sheet_ops,
            font=sheet_ops.table_font,
            bg=bg,
            fg=fg,
            select_bg=select_bg,
            select_fg=select_fg,
        )
        self.replace_tktext.reset(
            menu_kwargs=menu_kwargs,
            sheet_ops=sheet_ops,
            font=sheet_ops.table_font,
            bg=bg,
            fg=fg,
            select_bg=select_bg,
            select_fg=select_fg,
        )
        for widget in (
            self.find_previous_arrow,
            self.find_next_arrow,
            self.in_selection,
            self.close,
            self.toggle_replace,
            self.replace_next,
            self.replace_all,
        ):
            widget.config(
                font=sheet_ops.table_font,
                bg=bg,
                fg=fg,
                highlightbackground=bg,
                highlightcolor=fg,
            )
        if self.find_in_selection:
            self.enter_label(self.in_selection)
        self.config(
            background=bg,
            highlightbackground=border_color,
            highlightcolor=border_color,
            highlightthickness=1,
        )
        self.separator.config(background=grid_color)

        for b in sheet_ops.find_bindings:
            recursive_bind(self, b, self.focus_find)
        for b in sheet_ops.toggle_replace_bindings:
            recursive_bind(self, b, self.focus_replace)

        for b in sheet_ops.escape_bindings:
            recursive_bind(self, b, self.close_func)

        for b in sheet_ops.find_next_bindings:
            recursive_bind(self, b, self.find_next_func)
        for b in sheet_ops.find_previous_bindings:
            recursive_bind(self, b, self.find_prev_func)

    def start_tooltip_timer(self) -> None:
        self.tooltip_after_id = self.after(400, self.check_and_show_tooltip)

    def check_and_show_tooltip(self) -> None:
        """Check if the mouse position has changed and show tooltip if stationary."""
        if self.tooltip_widget is None:
            return
        current_x, current_y = get_mouse_coords(self.tooltip_widget)
        if current_x < 0 or current_y < 0:  # Mouse outside window
            return
        # Allow 2-pixel tolerance for minor movements
        if abs(current_x - self.tooltip_last_x) <= 2 and abs(current_y - self.tooltip_last_y) <= 2:
            self.show_tooltip(self.tooltip_widget)
        else:
            self.tooltip_last_x = current_x
            self.tooltip_last_y = current_y
            self.tooltip_after_id = self.after(400, self.check_and_show_tooltip)

    def show_tooltip(self, widget: tk.Misc) -> None:
        """Show the tooltip at the specified position."""
        bg = self.bg if self.bg is not None else "white"
        fg = self.fg if self.fg is not None else "black"
        self.tooltip = LabelTooltip(self, widget.tooltip_text, bg, fg)
        # Use current mouse position instead of recorded position
        self.tooltip.deiconify()
        show_x = max(0, self.winfo_toplevel().winfo_pointerx() - self.tooltip.winfo_width() - 5)
        show_y = self.winfo_toplevel().winfo_pointery()
        self.tooltip.wm_geometry(f"+{show_x}+{show_y - 10}")

    def cancel_tooltip(self):
        """Cancel any scheduled tooltip."""
        if self.tooltip_after_id is not None:
            self.after_cancel(self.tooltip_after_id)
            self.tooltip_after_id = None

    def hide_tooltip(self):
        """Hide the tooltip."""
        if self.tooltip is not None:
            self.tooltip.destroy()
            self.tooltip = None


def replacer(find: str, replace: str, current: str) -> Callable[[re.Match], str]:
    """Create a replacement function for re.sub with special empty string handling."""

    def _replacer(match: re.Match) -> str:
        if find:
            return replace
        else:
            if len(current) == 0:
                return replace
            else:
                return match.group(0)

    return _replacer


def get_mouse_coords(widget: tk.Misc) -> tuple[int, int]:
    # Get absolute mouse coordinates (relative to screen)
    mouse_x = widget.winfo_pointerx()
    mouse_y = widget.winfo_pointery()

    # Get widget's position relative to the screen
    widget_x = widget.winfo_rootx()
    widget_y = widget.winfo_rooty()

    # Calculate coordinates relative to the widget
    relative_x = mouse_x - widget_x
    relative_y = mouse_y - widget_y

    return relative_x, relative_y
