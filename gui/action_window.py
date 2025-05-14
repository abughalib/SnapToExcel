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
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollincrement="1")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

    def show(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        last_five_actions = list(self.storage.get_stack())[-5:]

        for action in reversed(last_five_actions):
            action_type: str = ""

            match action.action_type:
                case ACTION_TYPE.INSERT_IMAGE:
                    image_data = io.BytesIO(action.payload[0].getvalue())
                    img = Image.open(image_data)

                    # More Size support can be added here
                    if action.payload[1] == 1080:
                        img.thumbnail(IMAGE_THUMBNAIL_SIZE_1080P)
                    else:
                        img.thumbnail(IMAGE_THUMBNAIL_SIZE_1200P)

                    photo = ImageTk.PhotoImage(img)
                    label = tk.Label(self.scrollable_frame, image=photo)
                    label.image = photo
                    label.pack()

                case ACTION_TYPE.CHANGE_SHEET:
                    label = tk.Label(self.scrollable_frame, text="Change Sheet")
                    label.pack()

                case ACTION_TYPE.UNDO_LAST:
                    label = tk.Label(self.scrollable_frame, text="Undo Last Action")
                    label.pack()

                case ACTION_TYPE.INSERT_QUERY_RES:
                    text_widget = scrolledtext.ScrolledText(
                        self.scrollable_frame, height=10
                    )
                    text_widget.insert(tk.END, str(action.payload))
                    text_widget.config(state="disabled")
                    text_widget.pack()

            action_label = tk.Label(self.scrollable_frame, text=action_type)
            action_label.pack()
