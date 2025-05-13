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
        if SHORTCUTKEY_CONFIG in self.config[KEY_SECTION]:
            try:
                str_key_code = "".join(
                    [
                        ch
                        for ch in self.config[KEY_SECTION][SHORTCUTKEY_CONFIG]
                        if ch.isdigit()
                    ]
                )
                # TODO Fix numeric key shortcut
                logging.log(logging.INFO, f"Found Stored Shortcut Key: {str_key_code}")
                if str_key_code.isnumeric():
                    key_code = KeyCode.from_vk(int(str_key_code))
                    return Key(key_code)
                if str_key_code.isalpha():
                    key_code = KeyCode.from_char(str_key_code)
                    return Key(key_code)
            except Exception as e:
                logging.log(logging.ERROR, f"Failed to get SHORTCUT KEY")
                return DEFAULT_SHORTCUT_KEY

        logging.log(logging.INFO, "Returning DEFAULT_SHORTCUT_KEY")
        return DEFAULT_SHORTCUT_KEY

    def get_shortcut_key(self, key: Key | KeyCode):
        if type(key) == Key:
            logging.log(logging.INFO, f"Setting Shortcut Key to: {str(key.value)}")
            self.config[KEY_SECTION][SHORTCUTKEY_CONFIG] = str(key.value)
        else:
            logging.log(logging.INFO, f"Setting Shortcut Key to: {str(key.char)}")
            self.config[KEY_SECTION][SHORTCUTKEY_CONFIG] = str(key.char)

    def set_db_username(self, username: str):
        logging.log(logging.INFO, f"Setting DB Username to: {username}")
        self.config[DBAUTHENTICATIONSECTION][DBUSERNAME] = username

    def get_db_username(self) -> str:
        logging.log(logging.INFO, f"Getting DB Username")
        if DBUSERNAME in self.config[DBAUTHENTICATIONSECTION]:
            username = self.config[DBAUTHENTICATIONSECTION][DBUSERNAME]
            logging.log(logging.INFO, f"Found Stored DB Username: {username}")
            return username
        logging.log(logging.INFO, "Returning empty string for DB Username")
        return ""

    def get_db_password(self) -> str:
        logging.log(logging.INFO, f"Getting DB Password")
        if DBPASSWORD in self.config[DBAUTHENTICATIONSECTION]:
            password = self.config[DBAUTHENTICATIONSECTION][DBPASSWORD]
            logging.log(logging.INFO, f"Found Stored DB Password: {password}")
            return password
        logging.log(logging.INFO, "Returning empty string for DB Password")
        return ""

    def set_db_password(self, password: str):
        logging.log(logging.INFO, f"Setting DB Password to: {password}")
        self.config[DBAUTHENTICATIONSECTION][DBPASSWORD] = password

    def get_db_connection_url(self) -> str:
        logging.log(logging.INFO, f"Getting DB Connection URL")
        if DBCONNECTIONURL in self.config[DBAUTHENTICATIONSECTION]:
            url = self.config[DBAUTHENTICATIONSECTION][DBCONNECTIONURL]
            logging.log(logging.INFO, f"Found Stored DB Connection URL: {url}")
            return url
        logging.log(logging.INFO, "Returning default DB Connection URL")
        return DEFAULT_SERVICE_URL

    def set_db_connection_url(self, url: str):
        logging.log(logging.INFO, f"Setting DB Connection URL to: {url}")
        self.config[DBAUTHENTICATIONSECTION][DBCONNECTIONURL] = url

    def set_last_sql_queries(self, queries: str):
        logging.log(logging.INFO, f"Setting Last SQL Queries to: {queries}")
        self.config[SQLQUERIESSECTION][SQLQUERIES] = queries

    def get_last_sql_queries(self) -> str:
        logging.log(logging.INFO, f"Getting Last SQL Queries")
        if SQLQUERIES in self.config[SQLQUERIESSECTION]:
            queries = self.config[SQLQUERIESSECTION][SQLQUERIES]
            logging.log(logging.INFO, f"Found Stored Last SQL Queries: {queries}")
            return queries
        logging.log(logging.INFO, "Returning empty string for Last SQL Queries")
        return DEFAULTSQLQUERIES

    def get_last_path(self) -> str:
        logging.log(logging.INFO, f"Getting Last Path")
        if PATH_SECTION_CONFIG in self.config[PATH_SECTION]:
            path = self.config[PATH_SECTION][PATH_SECTION_CONFIG]
            logging.log(logging.INFO, f"Found Stored Last Path: {path}")
            return path
        logging.log(logging.INFO, "Returning current working directory for Last Path")
        return verify_dir_or_get_cur()

    def set_last_path(self, path: str):
        logging.log(logging.INFO, f"Setting Last Path to: {path}")
        self.config[PATH_SECTION][PATH_SECTION_CONFIG] = path

    def get_row_seperation(self) -> int:
        if ROW_SEPERATION in self.config[WORKBOOK_SECTION]:
            saved_row_seperation = self.config[WORKBOOK_SECTION][ROW_SEPERATION]
            logging.log(
                logging.INFO, f"Found Saved Row Seperation: {saved_row_seperation}"
            )
            if saved_row_seperation and saved_row_seperation.isnumeric():
                return int(saved_row_seperation)
            logging.log(logging.INFO, f"Invalid Row Seperation: {saved_row_seperation}")

        else:
            logging.log(
                logging.INFO,
                f"Row Seperation not found in config, using default: {DEFAULT_ROW_SEPERATION}",
            )
        return DEFAULT_ROW_SEPERATION

    def set_row_seperation(self, row_seperation: int):
        logging.log(logging.INFO, f"Setting Row Seperation to: {row_seperation}")
        self.config[WORKBOOK_SECTION][ROW_SEPERATION] = str(row_seperation)

    def get_workbook_start_position(self) -> Tuple[int, str]:
        logging.log(logging.INFO, f"Getting Workbook Start Position")
        if WORKBOOK_START_POSITION in self.config[WORKBOOK_SECTION]:
            start_position = self.config[WORKBOOK_SECTION][WORKBOOK_START_POSITION]
            logging.log(
                logging.INFO, f"Found Saved Workbook Start Position: {start_position}"
            )
        else:
            logging.log(logging.INFO, f"Workbook Start Position not found in config")

        return get_row_column(start_position)

    def set_workbook_start_position(self, start_position: str):
        logging.log(
            logging.INFO, f"Setting Workbook Start Position to: {start_position}"
        )
        self.config[WORKBOOK_SECTION][WORKBOOK_START_POSITION] = start_position

    def save_config(self):
        logging.log(logging.INFO, f"Saving config to: {get_config_path()}")
        with open(get_config_path(), "w") as configfile:
            self.config.write(configfile)
            logging.log(logging.INFO, f"Config saved successfully")
        logging.log(logging.ERROR, f"Config save failed")
