import logging
from pynput.keyboard import Key, Listener
from .screenshot import ScreenShot
from .export_sheet import XlxsSheet
from .models import ScreenshotMode
from gui import ui_constants
from datetime import datetime


class ShortcutKey:
    def __init__(
        self, xls: XlxsSheet, change_info_callback, shortcut_key=Key.f9
    ) -> None:
        logging.log(
            logging.INFO,
            f"Screenshot Initializing Screenshot Key: {shortcut_key} XlsxSheet is None: {xls == None}",
        )
        self.shortcut_key = shortcut_key
        self.flag = False
        self.listener = Listener(on_press=self.on_press)
        self.screenshot = ScreenShot(ScreenshotMode.FULLSCREEN)
        self.change_info_msg = change_info_callback
        self.xls = xls

    def on_press(self, key: Key) -> None:
        logging.log(logging.INFO, f"Detected Keypress: {self.shortcut_key}")
        if key == self.shortcut_key:
            logging.log(
                logging.INFO, f"Detected Shortcut Keypress: {self.shortcut_key}"
            )
            self.flag = True
            ss_image = self.screenshot.get_screenshot()
            if ss_image:
                self.change_info_msg(
                    f"Last screenshot taken at: {datetime.now()}", ui_constants.SUCCESS
                )
            else:
                self.change_info_msg(
                    f"Failed to take Screenshot at: {datetime.now()}",
                    ui_constants.ERROR,
                )
            logging.log(logging.INFO, f"Screenshot taken at: {datetime.now()}")
            self.xls.add_image_to_queue(ss_image)

    @staticmethod
    def change_shortcut_key() -> Key:
        pressed_key: Key = Key.f9
        logging.log(logging.INFO, "Changing shortcut Key")

        def on_press(key: Key) -> bool:
            nonlocal pressed_key
            pressed_key = key
            logging.log(logging.INFO, f"Detected KeyPresss: {key}")
            return False

        logging.log(logging.INFO, "Shortcut Key detect Listener Starting")
        with Listener(on_press=on_press) as listener:  # type: ignore
            listener.join()

        return pressed_key
