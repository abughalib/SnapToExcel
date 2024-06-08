import logging
import configparser
from typing import Tuple
from pynput.keyboard import KeyCode, Key
from .utils import get_config_path, verify_dir_or_get_cur, get_row_column

KEY_SECTION = "Keys"
SHORTCUTKEY_CONFIG = "SHORTCUTKEY"
DEFAULT_SHORTCUT_KEY = Key.f9

PATH_SECTION = "Paths"
PATH_SECTION_CONFIG = "LASTFOLDERPATH"

WORKBOOK_SECTION = "WORKBOOK_SECTION"
WORKBOOK_START_POSITION = "STARTPOSITION"
ROW_SEPERATION = "ROWSEPERATION"
DEFAULT_ROW_SEPERATION = 2

SQLQUERIESSECTION = "SQLQUERIES"
SQLQUERIES = "QUERY"
DEFAULTSQLQUERIES = ""

DBAUTHENTICATIONSECTION = "DBAUTHENTICATION"
DBUSERNAME = "DBUSERNAME"
DBPASSWORD = "DBPASSWORD"
DBCONNECTIONURL = "DBCONNECTIONURL"
DEFAULT_SERVICE_URL = "192.168.0.110:1234/SERVICE_NAME"


class SnapToExcelConfig:

    def __init__(self):
        logging.log(logging.INFO, f"Initializing SExcelConfig")
        self.config = configparser.ConfigParser()
        logging.log(logging.INFO, f"Reading config at: {get_config_path()}")
        self.config.read(get_config_path())
        self.config_section = self.config.sections()

        if KEY_SECTION not in self.config_section:
            logging.log(logging.INFO, f"Initializing KEY_SECTION = {{}}")
            self.config[KEY_SECTION] = {}

        if PATH_SECTION not in self.config_section:
            logging.log(logging.INFO, f"Initializing PATH_SECTION = {{}}")
            self.config[PATH_SECTION] = {}

        if WORKBOOK_SECTION not in self.config_section:
            logging.log(logging.INFO, f"Initializing WORKBOOK_SECTION = {{}}")
            self.config[WORKBOOK_SECTION] = {}

        if SQLQUERIESSECTION not in self.config_section:
            logging.log(logging.INFO, f"Initializing SQLQUERIESSECTION = {{}}")
            self.config[SQLQUERIESSECTION] = {}

        if DBAUTHENTICATIONSECTION not in self.config_section:
            logging.log(logging.INFO, f"Initializing DBAUTHENTICATIONSECTION = {{}}")
            self.config[DBAUTHENTICATIONSECTION] = {}

    def get_shortcut_key(self) -> Key:
        logging.log(logging.INFO, f"Get Shortcut Key")
