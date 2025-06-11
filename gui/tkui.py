import os
import re
import logging
import random
import threading
from collections import deque

import webbrowser
import tkinter as tk
from tkinter.font import Font
from PIL import Image, ImageTk
from pynput.keyboard import Key
from tkinter import ttk, StringVar, IntVar
from tkinter import filedialog, messagebox

from features.config import SnapToExcelConfig
from features.export_sheet import XlxsSheet
from features.shortcut_detect import ShortcutKey
from gui.action_window import LastFiveActions
from gui.sql import execute_bulk_query, get_connection
from gui.tooltip import CreateToolTip
from gui.notifier import Notification
from features.models import InfoType
from gui.ui_constants import *
from features import utils


class App(tk.Tk):

    def __init__(self):
        super().__init__()

        # Window
        self.title(APP_NAME)
        self.minsize(DEFAULT_VIEWPORT_WIDTH, DEFAULT_VIEWPORT_HEIGHT)
        self.resizable(True, True)
        self.attributes("-alpha", APP_TRANSPARENCY)  # type: ignore
        self.iconbitmap(utils.get_assets_file_path(APP_LARGE_ICON_PATH))  # type: ignore

        # Style
        self.style_text_info = ttk.Style()
        self.style_text_warning = ttk.Style()
        self.style_text_error = ttk.Style()
        self.style_text_success = ttk.Style()
        self.style_text_info.configure(STYLE_INFO_LABEL, foreground="black")  # type: ignore
        self.style_text_warning.configure(STYLE_WARNING_LABEL, foreground="orange")  # type: ignore
        self.style_text_error.configure(STYLE_ERROR_LABEL, foreground="red")  # type: ignore
        self.style_text_success.configure(STYLE_SUCCESS_LABEL, foreground="green")  # type: ignore

        # Required Initialization
        self.saved_config = SnapToExcelConfig()
        self.file_name = StringVar(value=self.generate_name())
        self.folder_path = StringVar(value=self.saved_config.get_last_path())
        self.db_username = StringVar(value=self.saved_config.get_db_username())
        self.db_password = StringVar(value=self.saved_config.get_db_password())
        self.sql_queries = self.saved_config.get_last_sql_queries()
        self.shortcutKey = None
        self.threads: deque[threading.Thread] = deque()
        self.shortcutKeyButton = self.saved_config.get_shortcut_key()
        self.start_row, self.start_column = (
            self.saved_config.get_workbook_start_position()
        )
        self.connection_url = StringVar(value=self.saved_config.get_db_connection_url())
        self.start_cell_position = StringVar(
            value=f"{self.start_column}{self.start_row}"
        )
        self.row_seperation = IntVar(value=self.saved_config.get_row_seperation())
        self.previous_sheet_name = DEFAULT_SHEET_NAME
        self.sheet_name = StringVar(value=DEFAULT_SHEET_NAME)
        self.message = DEFAULT_INFO_MESSAGE
        self.xls = None

        # Icons
        self.execute_sql_icon = self._get_photo_or_none(SQL_EXECUTE_BUTTON_ICON)
        self.change_sheet_icon = self._get_photo_or_none(CHANGE_SHEET_BUTTON_ICON)
        self.undo_last_action_icon = self._get_photo_or_none(UNDO_BUTTON_ICON)
        self.start_button_icon = self._get_photo_or_none(
            START_BUTTON_ICON, BUTTON_ICON_WIDTH + 30, BUTTON_ICON_HEIGHT + 15
        )
        self.stop_button_icon = self._get_photo_or_none(
            STOP_BUTTON_ICON, BUTTON_ICON_WIDTH + 30, BUTTON_ICON_HEIGHT + 15
        )
        self.change_shortcut_icon = self._get_photo_or_none(CHANGE_SHORTCUT_ICON)

        # Menubar
        self.menubar = tk.Menu(self)
        self.config(menu=self.menubar)  # type: ignore
        self.help_menu = tk.Menu(self.menubar, tearoff=False)
        self.help_menu.add_command(
            label=HELP_MENU_VIDEO_LABEL, command=self._handle_help_video
        )
        self.help_menu.add_command(
            label=HELP_MENU_DOCUMENTATION_LABEL, command=self._handle_help_docs
        )
        self.help_menu.add_command(
            label=HELP_MENU_CHECK_UPDATE, command=self._handle_help_check_update
        )
        self.help_menu.add_command(
            label=HELP_MENU_ABOUT_LABEL, command=self._handle_help_about
        )
        self.menubar.add_cascade(label=HELP_MENU_LABEL, menu=self.help_menu)

        # FILE NAME
        self.file_name_frame = ttk.Frame(self, name=FILE_NAME_SOURCE)
        self.file_name_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.file_name_label = ttk.Label(
            self.file_name_frame,
            text=FILE_NAME_LABEL,
            width=MAX_LABEL_WIDTH,
            name=FILE_NAME_SOURCE + "_label",
        )

        self.file_name_label.pack(side=tk.LEFT)
        self.file_name_textbox = ttk.Entry(
            self.file_name_frame,
            width=DEFAULT_TEXT_FIELD_WIDTH,
            textvariable=self.file_name,
            name=FILE_NAME_SOURCE + "_textbox",
        )
        self.file_name_textbox.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # FOLDER LOCATION
        self.folder_location_frame = ttk.Frame(self)
        self.folder_location_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.folder_location_label = ttk.Label(
            self.folder_location_frame,
            text=FOLDER_LOCATION_LABEL,
            width=MAX_LABEL_WIDTH,
        )

        self.folder_location_label.pack(side=tk.LEFT)
        self.folder_location_textbox = ttk.Entry(
            self.folder_location_frame,
            width=DEFAULT_TEXT_FIELD_WIDTH,
            textvariable=self.folder_path,
        )
        self.folder_location_textbox.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.folder_location_button = ttk.Button(
            self.folder_location_textbox, text="Browse", command=self.browse_folder
        )

        self.folder_location_button.pack(
            side=tk.RIGHT, in_=self.folder_location_textbox
        )

        # START CELL POSITION
        self.start_cell_position_frame = ttk.Frame(self)
        self.start_cell_position_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.start_cell_label = ttk.Label(
            self.start_cell_position_frame,
            text=START_CELL_POSITION_LABEL,
            width=MAX_LABEL_WIDTH,
        )
        self.start_cell_label.pack(side=tk.LEFT)
        self.start_cell_textbox = ttk.Entry(
            self.start_cell_position_frame,
            width=DEFAULT_TEXT_FIELD_WIDTH,
            textvariable=self.start_cell_position,
        )
        self.start_cell_textbox.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # ROW SEPERATION
        self.row_seperation_label_frame = ttk.Frame(self)
        self.row_seperation_label_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.row_seperation_label = ttk.Label(
            self.row_seperation_label_frame,
            text=ROW_SEPERATION_LABEL,
            width=MAX_LABEL_WIDTH,
        )
        self.row_seperation_label.pack(side=tk.LEFT)
        self.row_seperation_textbox = ttk.Entry(
            self.row_seperation_label_frame,
            width=DEFAULT_TEXT_FIELD_WIDTH,
            textvariable=self.row_seperation,
        )
        self.row_seperation_textbox.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # SHEET NAME
        self.sheet_name_label_frame = ttk.Frame(self)

        self.sheet_name_label_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.sheet_name_label = ttk.Label(
            self.sheet_name_label_frame, text=SHEET_NAME_LABEL, width=MAX_LABEL_WIDTH
        )
        self.sheet_name_label.pack(side=tk.LEFT)
        self.sheet_name_textbox = ttk.Entry(
            self.sheet_name_label_frame,
            width=DEFAULT_TEXT_FIELD_WIDTH,
            textvariable=self.sheet_name,
        )
        self.sheet_name_textbox.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # DB Credentials Username
        self.db_username_label_frame = ttk.Frame(self)
        self.db_username_label_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.db_username_label = ttk.Label(
            self.db_username_label_frame,
            text=DB_USERNAME_LABEL,
            width=MAX_LABEL_WIDTH,
        )
        self.db_username_label.pack(side=tk.LEFT)
        self.db_username_entry = ttk.Entry(
            self.db_username_label_frame,
            width=DEFAULT_TEXT_FIELD_WIDTH,
            textvariable=self.db_username,
        )
        self.db_username_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # DB Credentials Password

        self.db_password_label_frame = ttk.Frame(self)
        self.db_password_label_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.db_password_label = ttk.Label(
            self.db_password_label_frame,
            text=DB_PASSWORD_LABEL,
            width=MAX_LABEL_WIDTH,
        )

        self.db_password_label.pack(side=tk.LEFT)
        self.db_password_entry = ttk.Entry(
            self.db_password_label_frame,
            width=DEFAULT_TEXT_FIELD_WIDTH,
            show="*",
            textvariable=self.db_password,
        )
        self.db_password_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # DB Credentials Connection URL
        self.connection_url_label_frame = ttk.Frame(self)
        self.connection_url_label_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.connection_url_label = ttk.Label(
            self.connection_url_label_frame,
            text=DB_CONNECTION_URL_LABEL,
            width=MAX_LABEL_WIDTH,
        )
        self.connection_url_label.pack(side=tk.LEFT)
        self.connection_url_entry = ttk.Entry(
            self.connection_url_label_frame,
            width=DEFAULT_TEXT_FIELD_WIDTH,
            textvariable=self.connection_url,
        )
        self.connection_url_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # SQL Queries
        self.sql_queries_entry_frame = ttk.Frame(border=5)
        self.sql_queries_entry_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.sql_queries_entry = tk.Text(
            self.sql_queries_entry_frame,
            width=DEFAULT_TEXT_AREA_WIDTH + 3,
            height=DEFAULT_TEXT_AREA_HEIGHT,
            wrap=tk.WORD,
            undo=True,
        )

        # Add right-click menu
        self.context_menu = tk.Menu(self.sql_queries_entry, tearoff=0)
        self.context_menu.add_command(
            label=CUT_COMMAND_LABEL,
            command=lambda: self.sql_queries_entry.event_generate(CUT_EVENT),
        )
        self.context_menu.add_command(
            label=COPY_COMMAND_LABEL,
            command=lambda: self.sql_queries_entry.event_generate(COPY_EVENT),
        )
        self.context_menu.add_command(
            label=PASTE_COMMAND_LABEL,
            command=lambda: self.sql_queries_entry.event_generate(PASTE_EVENT),
        )
        self.context_menu.add_separator()
        self.context_menu.add_command(
            label=SELECT_ALL_COMMAND_LABEL,
            command=lambda: self.sql_queries_entry.tag_add("sel", "1.0", "end"),
        )

        # Bind right click to show context menu
        self.sql_queries_entry.bind(
            RIGHT_CLICK_BUTTON, lambda e: self.context_menu.post(e.x_root, e.y_root)
        )

        # Standard key bindings
        self.sql_queries_entry.bind(
            CTRL_A_BUTTON, lambda e: self.sql_queries_entry.tag_add("sel", "1.0", "end")
        )
        self.sql_queries_entry.bind(
            CTRL_Z_BUTTON, lambda e: self.sql_queries_entry.edit_undo()
        )
        self.sql_queries_entry.bind(
            CTRL_Y_BUTTON, lambda e: self.sql_queries_entry.edit_redo()
        )

        self.sql_queries_entry.insert("1.0", self.sql_queries)
        self.on_key_release()
        self.sql_queries_entry.bind(KEY_RELEASE, lambda event: self.on_key_release())
        self.sql_queries_entry.bind(ENTER_BUTTON, lambda event: self.on_key_release())
        self.sql_queries_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # BUTTONS
        self.first_row_button_frame = ttk.Frame()
        self.first_row_button_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Start Button
        self.start_process_btn = ttk.Button(
            self.first_row_button_frame,
            text=START_BUTTON_LABEL,
            command=self._start_process,
        )
        self.start_process_btn.config(image=self.start_button_icon)  # type: ignore
        self.start_process_btn.image = self.start_button_icon  # type: ignore
        self.start_process_btn.pack(side=tk.LEFT, expand=True)

        # Execute Query Button
        self.sql_queries_btn = ttk.Button(
            self.first_row_button_frame,
            text=SQL_QUERY_EXECUTE_BUTTON_LABEL,
            command=self._execute_sql,
        )

        self.sql_queries_btn.image = self.execute_sql_icon  # type: ignore
        self.sql_queries_btn.config(image=self.execute_sql_icon)  # type: ignore
        self.sql_queries_btn.pack(side=tk.LEFT, expand=True)

        # Change Sheet Button
        self.change_sheet_btn = ttk.Button(
            self.first_row_button_frame,
            text=CHANGE_SHEET_BUTTON_LABEL,
            command=self._change_sheet,
        )
        self.change_sheet_btn.image = self.change_sheet_icon  # type: ignore
        self.change_sheet_btn.config(image=self.change_sheet_icon)  # type: ignore
        self.change_sheet_btn.pack(side=tk.LEFT, expand=True)

        # Undo Button

        self.undo_action_btn = ttk.Button(
            self.first_row_button_frame,
            text=UNDO_LAST_ACTION_LABEL,
            command=self._undo_last_action,
        )
        self.undo_action_btn.config(image=self.undo_last_action_icon)  # type: ignore

        self.undo_action_btn.image = self.undo_last_action_icon  # type: ignore
        self.undo_action_btn.pack(side=tk.LEFT, expand=True)

        # InfoType.INFO Text
        self.info_text_frame = ttk.Frame()
        self.info_text_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.info_text_label = ttk.Label(self.info_text_frame, text=INFO_TEXT_LABEL)
        self.info_text_label.pack(side=tk.LEFT)

        # Change Shortcut Key
        self.change_shortcut_btn_frame = ttk.Frame()
        self.change_shortcut_btn_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.change_shortcut_btn = ttk.Button(
            self.change_shortcut_btn_frame,
            text=CHANGE_SHORTCUT_KEY_BUTTON_LABEL,
            image=self._get_photo_or_none(CHANGE_SHEET_BUTTON_ICON),  # type: ignore
            command=self._change_shortcut_key,
        )
        self.change_shortcut_btn.config(image=self.change_shortcut_icon)  # type: ignore
        self.change_shortcut_btn.pack(side=tk.LEFT)
        self.change_shortcut_btn.image = self.change_shortcut_icon  # type: ignore
        self.show_action_window_frame = ttk.Frame()
        self.show_action_window_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.show_action_window_frame_btn = ttk.Button(
            self.show_action_window_frame,
            text=SHOW_LAST_ACTION_BUTTON_LABEL,
            command=self._open_last_action_window,
        )
        self.show_action_window_frame_btn.pack(side=tk.LEFT)

        # Current Shortcut Key
        self.current_shortcut_key = ttk.Label(
            self.change_shortcut_btn_frame,
            text=f"Current Shortcut Key: {self.shortcutKeyButton}",
        )
        self.current_shortcut_key.pack(side=tk.LEFT)
        self.copyright_info = ttk.Label(self, text=COPYRIGHT_TEXT)
        self.copyright_info_font = Font(weight="bold", size=10)
        self.copyright_info.configure(font=self.copyright_info_font)
        self.copyright_info.pack(side=tk.LEFT)
        self.protocol("WM_DELETE_WINDOW", self.on_exit_gui)

        # ToolTips
        CreateToolTip(self.sql_queries_btn, SQL_QUERY_EXECUTE_BUTTON_LABEL)
        CreateToolTip(self.start_process_btn, START_BUTTON_LABEL)
        CreateToolTip(self.change_sheet_btn, CHANGE_SHEET_BUTTON_LABEL)
        CreateToolTip(self.undo_action_btn, UNDO_LAST_ACTION_LABEL)
        CreateToolTip(self.change_shortcut_btn, CHANGE_SHORTCUT_KEY_BUTTON_LABEL)

        # Extra
        self.notifier = Notification()

    def _open_last_action_window(self):

        if self.xls and self.xls.storage:
            last_action = self.xls.storage.get_last()
            if last_action:
                win = LastFiveActions(self, self.xls.storage)
                th = threading.Thread(target=win.show)
                th.start()
            else:
                self.change_info_text("Nothing to show")
        else:
            self.change_info_text("No Action taken yet...")

    def _change_sheet(self):

        if self.xls and self.sheet_name.get():
            if self.xls.sheet_name == self.sheet_name.get():
                logging.log(
                    logging.INFO,
                    f"Sheet Name same as previous",
                )
                self.change_info_text(f"Sheet Name same as Previous", InfoType.WARNING)
            else:
                logging.log(
                    logging.INFO,
                    f"Change Sheet Name from: {self.previous_sheet_name} to {self.sheet_name.get()}",
                )
                self.xls.change_sheet(self.sheet_name.get())
                self.change_sheet_btn.config(text=CHANGE_SHEET_BUTTON_LABEL_CHANGED)
                self.change_info_text(
                    f"Changed Sheet Name from: {self.previous_sheet_name} to {self.sheet_name.get()}",
                    InfoType.SUCCESS,
                )
                self.change_sheet_btn.config(text=CHANGE_SHEET_BUTTON_LABEL)
                self.previous_sheet_name = self.sheet_name.get()

        elif self.sheet_name.get():
            logging.log(
                logging.INFO,
                f"Change sheet Name of uninitialized Xls Sheet",
            )

            self.previous_sheet_name = self.sheet_name.get()
            self.change_sheet_btn.config(text=CHANGE_SHEET_BUTTON_LABEL_CHANGED)
            self.change_info_text(
                f"Changed First Sheet Name to {self.sheet_name.get()}", InfoType.SUCCESS
            )
            self.change_sheet_btn.config(text=CHANGE_SHEET_BUTTON_LABEL)
        else:

            logging.log(
                logging.ERROR,
                f"No Value to Change",
            )

            self.change_sheet_btn.config(text=CHANGE_SHEET_BUTTON_LABEL_FAILED)
            self.change_info_text(f"Cannot have empty sheet name", InfoType.WARNING)
            self.change_sheet_btn.config(text=CHANGE_SHEET_BUTTON_LABEL)

    def _set_row_seperation(self):
        new_value: str = self.row_seperation_textbox.get()
        if new_value and new_value.isnumeric():
            logging.log(
                logging.INFO,
                f"Changing row seperation from {self.row_seperation} to {new_value}",
            )

            if self.xls:
                self.row_seperation = int(new_value)
                self.xls.change_row_seperation(int(new_value))
                logging.log(
                    logging.INFO,
                    f"Row Seperation value update to: {new_value}",
                )
                self.change_info_text(
                    f"Updated row seperation value to: {new_value}", InfoType.SUCCESS
                )

            else:
                self.row_seperation = int(new_value)
                logging.log(
                    logging.INFO,
                    f"Set row seperation value to: {self.row_seperation}",
                )
                self.change_info_text(
                    f"Changed row seperation value to: {new_value}", InfoType.SUCCESS
                )
        else:
            logging.log(
                logging.ERROR,
                f"Invalid row seperation value: {new_value}",
            )
            self.change_info_text(
                f"Invalid row seperation value: {new_value}", InfoType.ERROR
            )

    def _set_start_position(self):

        new_value: str = self.start_cell_position.get()
        if new_value and utils.is_valid_excel_column(new_value):
            logging.log(
                logging.INFO,
                f"Setting Start Position {self.start_column} to {new_value}",
            )

            if self.xls:
                self.start_row, self.start_column = utils.get_row_column(new_value)
                self.xls.change_position(self.start_row, self.start_column)
                logging.log(
                    logging.INFO,
                    f"Updated start Position to: {new_value}",
                )
                self.change_info_text(
                    f"Updated Start Position to: {new_value}", InfoType.SUCCESS
                )
            else:
                self.start_row, self.start_column = utils.get_row_column(new_value)
                logging.log(
                    logging.INFO,
                    f"Set Start Position to: {self.row_seperation}",
                )
                self.change_info_text(
                    f"Set Start Position to: {new_value}", InfoType.SUCCESS
                )

        else:
            logging.log(
                logging.ERROR,
                f"Failed to set Start Position to: {new_value}",
            )
            self.change_info_text(
                f"Failed to set Start Position to: {new_value}", InfoType.ERROR
            )

    def _start_process(self):

        logging.log(logging.INFO, f"Start Process")
        if self.start_process_btn.cget("text") == START_BUTTON_LABEL:
            excel_save_folder = utils.verify_dir_or_get_cur(self.folder_path.get())
            excel_save_filename = utils.clean_xls_file_name(self.file_name.get())

            if utils.check_file_exists(
                utils.join_path(excel_save_folder, excel_save_filename)
            ):
                self.change_info_text(
                    "File Already Exists, Change File Name", InfoType.WARNING
                )

            else:
                self._set_start_position()
                self.xls = XlxsSheet(
                    excel_save_folder,
                    excel_save_filename,
                    self.start_row,
                    self.start_column,
                    self.sheet_name.get(),
                    self.row_seperation.get(),  # type: ignore
                )

                self.shortcutKey = ShortcutKey(
                    self.xls, self.change_info_text, self.shortcutKeyButton
                )
                self.start_process_btn.config(
                    text=STOP_BUTTON_LABEL,
                    image=self.stop_button_icon,  # type: ignore
                )
                self.start_process_btn.image = self.stop_button_icon  # type: ignore self.shortcutkey.listener.start()
                self.shortcutKey.listener.start()

                logging.log(
                    logging.INFO,
                    f"Started Listining to Shortcut Key: {self.shortcutKey.shortcut_key}",
                )
                self.change_info_text(
                    "Started Listening to Shortcut Key", InfoType.INFO
                )
        else:
            th = threading.Thread(target=self._save_sheet)
            self.threads.append(th)
            th.start()
            self.start_process_btn.config(text=START_BUTTON_LABEL, image=self.start_button_icon)  # type: ignore
            self.start_process_btn.image = self.start_button_icon  # type: ignore

            if self.shortcutKey:
                self.shortcutKey.listener.stop()
                print("Stopped Listining to Key")
                logging.log(
                    logging.INFO,
                    f"Stopped Listining to Key: {self.shortcutKey.shortcut_key}",
                )
                self.change_info_text(
                    "Stopped Listening to Shortcut Key", InfoType.INFO
                )

    def browse_folder(self):

        logging.log(logging.INFO, f"Browsing Folder for file location", exc_info=True)
        folder = filedialog.askdirectory()
        logging.log(logging.INFO, f"Selected Folder Dir: (folder)", exc_info=True)
        self.folder_path = StringVar(value=utils.verify_dir_or_get_cur(folder))
        self.folder_location_textbox.config(textvariable=self.folder_path)

    def _get_photo_or_none(
        self,
        asset_name: str,
        width: int = BUTTON_ICON_WIDTH,
        height: int = BUTTON_ICON_HEIGHT,
    ):

        if utils.check_file_exists(utils.get_assets_file_path(asset_name)):
            image = Image.open(utils.get_assets_file_path(asset_name))
            image = image.resize(size=(width, height))
            photoImage = ImageTk.PhotoImage(image)
            return photoImage

        return None

    def _save_sheet(self):
        logging.log(
            logging.INFO, f"Saving Excel file with Name: {self.file_name.get()}"
        )

        if self.xls and self.xls.save_sheet():

            xlsx_file_abs_path = os.path.abspath(self.xls.file_path)
            self.change_info_text(
                f"File Saved at: {xlsx_file_abs_path}", InfoType.SUCCESS
            )
            th = threading.Thread(
                target=self.notifier.send_notification_success,
                args=(
                    "File Saved",
                    f"At: {xlsx_file_abs_path}",
                    f"{xlsx_file_abs_path}",
                ),
            )
            self.threads.append(th)
            th.start()

        else:
            self.change_info_text(f"Saving Failed, Check Log", InfoType.ERROR)
            th = threading.Thread(
                target=self.notifier.send_notification_error,
                args=("Failed", f"Saving File Failed, Check Log", 5),
            )
            self.threads.append(th)
            th.start()

    def _change_shortcut_key(self):

        logging.log(logging.INFO, f"Changing shortcut key")
        pressed_key = ShortcutKey.change_shortcut_key()
        if pressed_key:
            logging.log(logging.INFO, f"Shortcut Key changed to: {pressed_key}")
            self.shortcutKeyButton = pressed_key
            self._change_current_shortcutkey_text(pressed_key)
            logging.log(logging.INFO, f"Destroying Dialog Box")

        else:
            logging.log(logging.ERROR, f"No KeyPress Detected")

    def _change_current_shortcutkey_text(self, new_key: Key):

        logging.log(logging.INFO, f"Changing Shortcut Key Text Value To: {new_key}")
        self.current_shortcut_key.config(text=f"Current Shortcut Key: {new_key}")

    def change_info_text(self, message: str, info_type: InfoType = InfoType.INFO):

        logging.log(logging.INFO, f"Info Text Changed to: {message}")
        style = STYLE_INFO_LABEL
        match info_type:
            case InfoType.WARNING:
                style = STYLE_WARNING_LABEL
            case InfoType.ERROR:
                style = STYLE_ERROR_LABEL
            case InfoType.SUCCESS:
                style = STYLE_SUCCESS_LABEL
            case InfoType.INFO:
                style = STYLE_INFO_LABEL

        self.info_text_label.config(
            text=f"{info_type.capitalize()}: {message}", style=style
        )

    def _undo_last_action(self):

        logging.log(logging.INFO, f"Undo Last Action")
        if self.xls:
            res = self.xls.undo_last_action()
            logging.log(logging.INFO, f"Last Action Removed")
            self.change_info_text(res, InfoType.SUCCESS)
        else:
            logging.log(logging.ERROR, f"xlsx not initialized yet")
            self.change_info_text("Process yet to start", InfoType.WARNING)

    def _save_config(self):

        logging.log(logging.INFO, f"Saving Configurations")
        self.saved_config.set_last_path(self.folder_path.get())
        self.saved_config.set_shortcut_key(self.shortcutKeyButton)
        self.saved_config.set_row_seperation(self.row_seperation.get())  # type: ignore
        self.saved_config.set_workbook_start_position(self.start_cell_position.get())
        self.saved_config.set_last_sql_queries(self.sql_queries)  # type: ignore
        self.saved_config.set_db_username(self.db_username.get())
        self.saved_config.set_db_password(self.db_password.get())
        self.saved_config.set_db_connection_url(self.connection_url.get())

        if self.xls and not self.xls.saved:
            logging.log(logging.INFO, f"Saving Excel Sheet")
            self.xls.save_sheet()
        self.saved_config.save_config()

    def _execute_sql(self):

        logging.log(logging.INFO, f"Execute SQL")
        self.sql_queries = self.sql_queries_entry.get(1.0, tk.END)
        striped_queries = utils.clean_sql_query(self.sql_queries)
        queries = [query for query in striped_queries.split(";") if query]
        connection_url = self.connection_url.get()

        if not utils.is_valid_connection_url(connection_url):

            self.change_info_text(
                f"Check Connection URL, It should be IP:PORT/SERVICE_NAME",
                InfoType.ERROR,
            )

            return

        hostname = connection_url.split(":")[0]
        port, service_name = connection_url.split(":")[1].split("/")
        conn_string = f"{self.db_username.get()}/{self.db_password.get()}@{hostname}:{port}/{service_name}"

        if self.xls:
            conn = get_connection(conn_string)
            if conn:
                res = execute_bulk_query(conn, queries)
                if res:
                    if self.xls:
                        self.xls.insert_list_values_queue(res)
                        total_rows_feched = sum([len(x) for x in res])
                        self.change_info_text(
                            f"Executed Query Sucessfully, Fetched: {total_rows_feched} rows",
                            InfoType.SUCCESS,
                        )
                        conn.close()
                else:
                    logging.log(
                        logging.ERROR, f"Query Execution Failed, {self.sql_queries}"
                    )
                    self.change_info_text(f"Query Execution Failed", InfoType.ERROR)
            else:
                logging.log(logging.ERROR, f"Connection Creation Failed: {conn_string}")
                self.change_info_text(f"Connection Creation Failed", InfoType.ERROR)
        else:
            logging.log(logging.ERROR, f"XLS not initialized yet")
            self.change_info_text("Process yet to start", InfoType.WARNING)

    def _handle_help_video(self):
        webbrowser.open_new(HELP_DOCS_VIDEO)

    def _handle_help_docs(self):
        webbrowser.open_new(HELP_DOCS_LINK)

    def _handle_help_check_update(self):
        webbrowser.open_new(UPDATE_LINK_DOCS)

    def _handle_help_about(self):
        ABOUT_MESSAGE = f"""SnapToExcel:
            
            \nA Python based tool to automatically create excel sheets from Screenshot and executing SQL queries.
            
            \nVersion: {APP_VERSION}\
            \nAuthor: {APP_AUTHOR}\
            \nEmail: {APP_AUTHOR_EMAIL}\
        """

        messagebox.showinfo(title=HELP_MENU_ABOUT_LABEL, message=ABOUT_MESSAGE)  # type: ignore

    def on_exit_gui(self):

        logging.log(logging.INFO, f"---Exiting App---")
        for th in self.threads:
            th.join()

        self._save_config()
        self.destroy()

    def configure_tags(self, tags: dict[str, str]):

        for tag, color in tags.items():
            self.sql_queries_entry.tag_delete(tag)
            self.sql_queries_entry.tag_config(tag, foreground=color)

    def on_key_release(self):

        th = threading.Thread(target=self._on_key_release)  # type: ignore
        th.start()

    def _on_key_release(self):

        text_widget = self.sql_queries_entry

        content = text_widget.get("1.0", tk.END)

        for tag in text_widget.tag_names():
            text_widget.tag_remove(tag, "1.0", tk.END)

        styles = {
            "comment": {
                "foreground": "#006400",
                "font": ("Consolas", 10, "italic"),
            },  # dark green
            "string": {"foreground": "#C7601A"},  # deep orange
            "keyword": {
                "foreground": "#00008B",
                "font": ("Consolas", 10, "bold"),
            },  # dark blue
            "function": {"foreground": "#008B8B"},  # dark cyan
            "operator": {"foreground": "#555555"},  # dark gray
            "number": {"foreground": "#004B23"},  # darker green
            "punctuation": {"foreground": "#333333"},  # very dark gray
            "identifier": {"foreground": "#000000"},  # pure black
        }
        for tag, cfg in styles.items():
            text_widget.tag_configure(tag, cfg)

        patterns = [
            (re.compile(r"--.*?$", re.MULTILINE), "comment"),
            (re.compile(r"/\*.*?\*/", re.DOTALL), "comment"),
            (re.compile(r"'(?:[^'\\]|\\.)*'"), "string"),
            (
                re.compile(
                    r"\b(ADD|ALL|ALTER|AND|ANY|AS|ASC|BACKUP|BETWEEN|BY|CASE|CHECK|COLUMN|CONSTRAINT|CREATE|DATABASE|DEFAULT|DELETE|DESC|DISTINCT|DROP|EXEC|EXISTS|FOREIGN|FROM|FULL|GROUP|HAVING|IN|INDEX|INNER|INSERT|INTO|IS|JOIN|KEY|LEFT|LIKE|LIMIT|NOT|NULL|ON|OR|ORDER|OUTER|PRIMARY|PROCEDURE|RIGHT|ROWNUM|SELECT|SET|TABLE|TOP|UNION|UNIQUE|UPDATE|VALUES|VIEW|WHERE)\b",
                    re.IGNORECASE,
                ),
                "keyword",
            ),
            (re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*(?=\()"), "function"),
            (re.compile(r"=|<>|<|>|<=|>=|\+|\-|\*|/|\|\||%"), "operator"),
            (re.compile(r"\b\d+\b"), "number"),
            (re.compile(r"[(),.;]"), "punctuation"),
            (re.compile(r"\b[A-Za-z_][A-Za-z0-9_]*\b"), "identifier"),
        ]

        for regex, tag in patterns:
            for match in regex.finditer(content):
                start, end = match.start(), match.end()
                start_index = text_widget.index(f"1.0+{start}c")
                end_index = text_widget.index(f"1.0+{end}c")
                text_widget.tag_add(tag, start_index, end_index)

    @staticmethod
    def generate_name():
        chars = "abcdefghijklmnopqrstuvwxyz@123456789"
        name = ""

        for _ in range(7):
            index = int(len(chars) * random.random())
            char = chars[index]
            name += char

        logging.log(logging.INFO, f"Generated Name: {name}")

        return name

    def startApp(self):

        self.style = ttk.Style()
        self.mainloop()
