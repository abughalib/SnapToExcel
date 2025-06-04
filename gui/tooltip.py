import tkinter.ttk as ttk
import tkinter as tk


class CreateToolTip(object):

    def __init__(self, widget: ttk.Button, text: str = "Widget Info"):
        self.wraplength = 180  # pixels
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)  # type: ignore
        self.widget.bind("<Leave>", self.leave)  # type: ignore
        self.widget.bind("<ButtonPress>", self.leave)  # type: ignore
        self.id = None
        self.tw = None

    def enter(self, _: tk.EventType) -> None:
        self.showtip()

    def leave(self, _: tk.EventType):
        self.hidetip()

    def showtip(self):
        x = y = 0
        position: tuple[int, int, int, int] | None = self.widget.bbox()

        if position:
            x += self.widget.winfo_rootx() + 25
            y += self.widget.winfo_rooty() + 20

        self.tw = tk.Toplevel(self.widget)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(
            self.tw,
            text=self.text,
            justify="left",
            background="#ffffff",
            relief="solid",
            borderwidth=1,
            wraplength=self.wraplength,
        )
        label.pack(ipadx=1)

    def hidetip(self):
        if self.tw:
            self.tw.destroy()
            self.tw = None
