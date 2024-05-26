from __future__ import annotations

import tkinter as tk

from .vars import rc_binding


class TopLeftRectangle(tk.Canvas):
    def __init__(self, *args, **kwargs) -> None:
        tk.Canvas.__init__(
            self,
            kwargs["parent"],
            background=kwargs["parent"].ops.top_left_bg,
            highlightthickness=0,
        )
        self.PAR = kwargs["parent"]
        self.MT = kwargs["main_canvas"]
        self.RI = kwargs["row_index_canvas"]
        self.CH = kwargs["header_canvas"]
        try:
            self.config(
                width=self.RI.current_width,
                height=self.CH.current_height,
            )
        except Exception:
            return
        self.extra_motion_func = None
        self.extra_b1_press_func = None
        self.extra_b1_motion_func = None
        self.extra_b1_release_func = None
        self.extra_double_b1_func = None
        self.extra_rc_func = None
        self.MT.TL = self
        self.RI.TL = self
        self.CH.TL = self
        w = self.RI.current_width
        h = self.CH.current_height
        self.rw_box = self.create_rectangle(
            0,
            h - 5,
            w,
            h,
            fill=self.PAR.ops.top_left_fg,
            outline="",
            tag="rw",
            state="normal" if self.RI.width_resizing_enabled else "hidden",
        )
        self.rh_box = self.create_rectangle(
            w - 5,
            0,
            w,
            h,
            fill=self.PAR.ops.top_left_fg,
            outline="",
            tag="rh",
            state="normal" if self.CH.height_resizing_enabled else "hidden",
        )
        self.select_all_box = self.create_rectangle(
            0,
            0,
            w - 5,
            h - 5,
            fill=self.PAR.ops.top_left_bg,
            outline="",
            tag="sa",
            state="normal" if self.MT.select_all_enabled else "hidden",
        )
        self.select_all_tri = self.create_polygon(
            w - 7,
            h - 7 - 10,
            w - 7,
            h - 7,
            w - 7 - 10,
            h - 7,
            fill=self.PAR.ops.top_left_fg,
            outline="",
            tag="sa",
            state="normal" if self.MT.select_all_enabled else "hidden",
        )
        self.tag_bind("rw", "<Enter>", self.rw_enter)
        self.tag_bind("rh", "<Enter>", self.rh_enter)
        self.tag_bind("sa", "<Enter>", self.sa_enter)
        self.tag_bind("rw", "<Leave>", self.rw_leave)
        self.tag_bind("rh", "<Leave>", self.rh_leave)
        self.tag_bind("sa", "<Leave>", self.sa_leave)
        self.bind("<Motion>", self.mouse_motion)
        self.bind("<ButtonPress-1>", self.b1_press)
        self.bind("<B1-Motion>", self.b1_motion)
        self.bind("<ButtonRelease-1>", self.b1_release)
        self.bind("<Double-Button-1>", self.double_b1)
        self.bind(rc_binding, self.rc)

    def redraw(self) -> None:
        self.itemconfig("rw", fill=self.PAR.ops.top_left_fg)
        self.itemconfig("rh", fill=self.PAR.ops.top_left_fg)
        self.itemconfig(
            self.select_all_tri,
            fill=self.PAR.ops.top_left_fg,
        )

    def rw_state(self, state: str = "normal") -> None:
        self.itemconfig("rw", state=state)

    def rh_state(self, state: str = "normal") -> None:
        self.itemconfig("rh", state=state)

    def sa_state(self, state: str = "normal") -> None:
        self.itemconfig("sa", state=state)

    def rw_enter(self, event: object = None) -> None:
        if self.RI.width_resizing_enabled:
            self.itemconfig(
                "rw",
                fill=self.PAR.ops.top_left_fg_highlight,
            )

    def sa_enter(self, event: object = None) -> None:
        if self.MT.select_all_enabled:
            self.itemconfig(
                self.select_all_tri,
                fill=self.PAR.ops.top_left_fg_highlight,
            )

    def rh_enter(self, event: object = None) -> None:
        if self.CH.height_resizing_enabled:
            self.itemconfig(
                "rh",
                fill=self.PAR.ops.top_left_fg_highlight,
            )

    def rw_leave(self, event: object = None) -> None:
        self.itemconfig("rw", fill=self.PAR.ops.top_left_fg)

    def rh_leave(self, event: object = None) -> None:
        self.itemconfig("rh", fill=self.PAR.ops.top_left_fg)

    def sa_leave(self, event: object = None) -> None:
        self.itemconfig(
            self.select_all_tri,
            fill=self.PAR.ops.top_left_fg,
        )

    def basic_bindings(self, enable: bool = True) -> None:
        if enable:
            self.bind("<Motion>", self.mouse_motion)
            self.bind("<ButtonPress-1>", self.b1_press)
            self.bind("<B1-Motion>", self.b1_motion)
            self.bind("<ButtonRelease-1>", self.b1_release)
            self.bind("<Double-Button-1>", self.double_b1)
            self.bind(rc_binding, self.rc)
        else:
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")
            self.unbind(rc_binding)

    def set_dimensions(
        self,
        new_w: None | int = None,
        new_h: None | int = None,
        recreate_selection_boxes: bool = True,
    ) -> None:
        try:
            if new_h is None:
                h = self.winfo_height()
            if new_w is None:
                w = self.winfo_width()
            if new_w:
                self.config(width=new_w)
                w = new_w
            if new_h:
                self.config(height=new_h)
                h = new_h
        except Exception:
            return
        self.coords(self.rw_box, 0, h - 5, w, h)
        self.coords(self.rh_box, w - 5, 0, w, h)
        self.coords(
            self.select_all_tri,
            w - 7,
            h - 7 - 10,
            w - 7,
            h - 7,
            w - 7 - 10,
            h - 7,
        )
        self.coords(self.select_all_box, 0, 0, w - 5, h - 5)
        if recreate_selection_boxes:
            self.MT.recreate_all_selection_boxes()

    def mouse_motion(self, event: object = None) -> None:
        self.MT.reset_mouse_motion_creations()
        if self.extra_motion_func is not None:
            self.extra_motion_func(event)

    def b1_press(self, event: object = None) -> None:
        self.focus_set()
        rect = self.find_overlapping(event.x, event.y, event.x, event.y)
        if not rect or rect[0] in (
            self.select_all_box,
            self.select_all_tri,
        ):
            if self.MT.select_all_enabled:
                self.MT.deselect("all")
                self.MT.select_all()
            else:
                self.MT.deselect("all")
        elif rect[0] == self.rw_box:
            if self.RI.width_resizing_enabled:
                self.RI.set_width(
                    self.PAR.ops.default_row_index_width,
                    set_TL=True,
                )
        elif rect[0] == self.rh_box:
            if self.CH.height_resizing_enabled:
                self.CH.set_height(
                    self.MT.get_default_header_height(),
                    set_TL=True,
                )
        self.MT.main_table_redraw_grid_and_text(
            redraw_header=True,
            redraw_row_index=True,
        )
        if self.extra_b1_press_func is not None:
            self.extra_b1_press_func(event)

    def b1_motion(self, event: object = None) -> None:
        self.focus_set()
        if self.extra_b1_motion_func is not None:
            self.extra_b1_motion_func(event)

    def b1_release(self, event: object = None) -> None:
        self.focus_set()
        if self.extra_b1_release_func is not None:
            self.extra_b1_release_func(event)

    def double_b1(self, event: object = None) -> None:
        self.focus_set()
        if self.extra_double_b1_func is not None:
            self.extra_double_b1_func(event)

    def rc(self, event: object = None) -> None:
        self.focus_set()
        if self.extra_rc_func is not None:
            self.extra_rc_func(event)
