import io

from PIL import ImageTk, Image
from tkinter import scrolledtext
import tkinter as tk

from features.storage import Storage
from features.actions import ACTION_TYPE
from gui.ui_constants import *


class LastFiveActions:
    def __init__(self, parent: tk.Tk, storage: Storage):
        self.parent = parent
        self.storage = storage
        self.child_window = tk.Toplevel(self.parent)
        self.child_window.title(LAST_ACTION_TITLE)
        self.create_scrollable_window()

    def create_scrollable_window(self):
        self.canvas = tk.Canvas(self.child_window)
        self.scrollbar = tk.Scrollbar(
            self.child_window, orient="vertical", command=self.canvas.yview
        )

        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0))