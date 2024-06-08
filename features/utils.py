import re
import os
import uuid
import logging
import datetime
import platform
from typing import Tuple
from .constants import *
from zipfile import ZipFile


def get_current_dir():
    logging.log(logging.INFO, f"Current DIR: {os.curdir}", exc_info=True)
    return os.curdir


def get_image_path():

    screen_shot_path = os.environ.get("SCREENSHOT_IMAGE_PATH")
    logging.log(logging.INFO, f"Image Path: {screen_shot_path}", exc_info=True)
    if screen_shot_path:
        return screen_shot_path

    logging.log(logging.ERROR, f"Image path not found in Environment Variable")

    return os.path.curdir

def get_image_file_name():
    file_name = clean_file_name(f"{DEFAULT_APP_NAME}-{datetime.datetime.now()}.png")
    logging.log(logging.INFO, f"Image File Name: {file_name}", exc_info=True)
    