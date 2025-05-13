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


def clean_file_name(file_name: str) -> str:
    logging.log(logging.INFO, f"Cleaning File Name: {file_name}", exc_info=True)
    file_name = re.sub(r"[\\/:*?\"<>|]", "_", file_name)
    file_name = re.sub(r"\s+", "_", file_name)
    file_name = re.sub(r"[^a-zA-Z0-9_.-]", "", file_name)
    logging.log(logging.INFO, f"Cleaned File Name: {file_name}", exc_info=True)

    return file_name


def get_image_file_name():
    file_name = clean_file_name(f"{DEFAULT_APP_NAME}-{datetime.datetime.now()}.png")
    logging.log(logging.INFO, f"Image File Name: {file_name}", exc_info=True)

    return file_name


def get_user_dir():
    logging.log(logging.INFO, f"User DIR: {os.path.expanduser('~')}", exc_info=True)
    return os.path.expanduser("~")


def get_image_full_path():
    full_path = os.path.join(get_image_path(), get_image_file_name())
    logging.log(logging.INFO, f"Image Full Path: {full_path}", exc_info=True)

    return full_path


def get_excel_file_path():
    excel_file_path = os.environ.get("SCREENSHOT_FILE_PATH")
    logging.log(logging.INFO, f"Excel File Path: {excel_file_path}", exc_info=True)
    if excel_file_path:
        logging.log(
            logging.INFO,
            f"Excel File Path: {excel_file_path} found in Environment Variable",
            exc_info=True,
        )
        return excel_file_path
    logging.log(logging.ERROR, f"Excel file path not found in Environment Variable")
    return os.path.curdir


def join_path(path: str, file_name: str) -> str:
    logging.log(logging.INFO, f"Joining Path: {path} and File Name: {file_name}")
    if not os.path.exists(path):
        logging.log(
            logging.INFO,
            f"Path: {path} does not exist, creating it",
            exc_info=True,
        )
        os.makedirs(path)

    full_path = os.path.join(path, file_name)
    logging.log(logging.INFO, f"Full Path: {full_path}", exc_info=True)

    return full_path


def verify_dir(path: str):
    logging.log(logging.INFO, f"Verifying Directory: {path}", exc_info=True)
    temp_file = os.path.join(path, str(uuid.uuid4()))
    try:
        with open(temp_file, "w") as f:
            f.write("test")
        os.remove(temp_file)
        logging.log(logging.INFO, f"Directory: {path} is valid", exc_info=True)
        return True
    except Exception as e:
        logging.log(
            logging.ERROR, f"Directory: {path} is not valid: {e}", exc_info=True
        )
        return False


def verify_file(path: str):
    logging.log(logging.INFO, f"Verifying File: {path}", exc_info=True)

    if os.path.exists(path):
        try:
            with open(path, "r") as f:
                f.read()
            logging.log(logging.INFO, f"File: {path} is valid", exc_info=True)
            return True
        except PermissionError as e:
            logging.log(logging.ERROR, f"File: {path} is not valid: {e}", exc_info=True)
            return False

        except Exception as e:
            logging.log(logging.ERROR, f"File: {path} is not valid: {e}", exc_info=True)
            return False

    logging.log(logging.ERROR, f"File: {path} does not exist", exc_info=True)
    return False


def verify_dir_or_get_cur(path: str):
    logging.log(logging.INFO, f"Verifying Directory: {path}", exc_info=True)
    if verify_dir(path):
        logging.log(logging.INFO, f"Directory: {path} is valid", exc_info=True)
        return path

    logging.log(
        logging.INFO,
        f"Directory: {path} is not valid, returning current working directory",
        exc_info=True,
    )

    return os.path.curdir


def check_file_exists(file_path: str) -> bool:
    logging.log(logging.INFO, f"Checking if file exists: {file_path}", exc_info=True)
    if os.path.exists(file_path):
        logging.log(logging.INFO, f"File exists: {file_path}", exc_info=True)
        return True
    else:
        logging.log(logging.ERROR, f"File does not exist: {file_path}", exc_info=True)
        return False


def clean_xls_file_name(file_name: str) -> str:
    logging.log(logging.INFO, f"Cleaning XLS File Name: {file_name}", exc_info=True)

    if len(file_name) < 6:
        return clean_file_name(file_name) + ".xlsx"
    name_split = file_name.split(".xlsx")

    if name_split and name_split[-1] != "xlsx":
        logging.log(
            logging.INFO,
            f"File Name: {file_name} does not end with .xlsx, cleaning it",
            exc_info=True,
        )
        file_name = name_split[0] + ".xlsx"

        return clean_file_name(file_name) + ".xlsx"

    return file_name


def get_log_path():
    log_final_path = os.path.join(
        get_user_dir(),
        clean_file_name(f"{DEFAULT_APP_NAME}-log-{datetime.datetime.now().date()}.log"),
    )

    return os.path.join(get_current_dir(), log_final_path)


def _config_folder() -> str:
    logging.log(logging.INFO, f"Config Folder: {DEFAULT_CONFIG_DIR}", exc_info=True)
    user_dir = get_user_dir()

    config_folder = os.path.join(user_dir, DEFAULT_CONFIG_DIR)

    if not os.path.exists(config_folder):
        logging.log(
            logging.INFO,
            f"Config Folder: {config_folder} does not exist, creating it",
            exc_info=True,
        )
        os.makedirs(config_folder)

    logging.log(logging.INFO, f"Config Folder: {config_folder}", exc_info=True)
    return config_folder


def get_config_path() -> str:
    logging.log(logging.INFO, f"Config Path: {DEFAULT_CONFIG_DIR}", exc_info=True)

    config_folder = _config_folder()

    config_file_name = DEFAULT_APP_NAME + "-config.ini"

    config_file_path = os.path.join(config_folder, config_file_name)

    logging.log(
        logging.INFO,
        f"Returning Config File Path: {config_file_path}",
        exc_info=True,
    )

    return config_file_path


def is_valid_excel_column(cell_pos: str):
    logging.log(logging.INFO, f"Validating Excel Column: {cell_pos}", exc_info=True)

    return re.match(r"^[A-Z]+\\d+$", cell_pos) is not None


def get_row_column(cell_pos: str) -> Tuple[int, str]:
    logging.log(
        logging.INFO,
        f"Getting Row and Column from Cell Position: {cell_pos}",
        exc_info=True,
    )

    if is_valid_excel_column(cell_pos):
        match = [x for x in re.split("(\\d+)|([A-Z]+)", cell_pos) if x]
        if len(match) == 2:
            row: int = int(match[1])
            column: str = match[0]
            logging.log(
                logging.INFO,
                f"Row: {row}, Column: {column} from Cell Position: {cell_pos}",
                exc_info=True,
            )

            return row, column

    logging.log(
        logging.ERROR,
        f"Invalid Cell Position: {cell_pos}, returning default values",
        exc_info=True,
    )

    return DEFAULT_START_ROW, DEFAULT_START_COLUMN


def is_valid_connection_url(connection_url: str) -> bool:

    if re.findall(ORACLE_CONNECTION_URL_VALIDATION_REGEX, connection_url):
        logging.log(
            logging.INFO,
            f"Valid Connection URL: {connection_url}",
            exc_info=True,
        )
        return True

    logging.log(
        logging.ERROR,
        f"Invalid Connection URL: {connection_url}",
        exc_info=True,
    )

    return False


def is_valid_excel_column_name(column_name: str) -> bool:
    logging.log(
        logging.INFO,
        f"Validating Excel Column Name: {column_name}",
        exc_info=True,
    )

    if re.match(r"^[A-Z]+$", column_name):
        logging.log(
            logging.INFO,
            f"Valid Column Name: {column_name}",
            exc_info=True,
        )
        return True

    logging.log(
        logging.ERROR,
        f"Invalid Column Name: {column_name}",
        exc_info=True,
    )

    return False


def get_asses_path():
    return join_path(get_current_dir(), "assets")


def get_assets_file_path(file_name: str):
    return os.path.abspath(join_path(get_asses_path(), file_name))


def _get_oracle_driver_path_for_os():
    driver_path = ""
    if platform.system() == "Windows":
        logging.log(
            logging.INFO,
            f"Windows OS detected, using driver path: {ORACLE_WINDOWS_DRIVER_PATH}",
            exc_info=True,
        )
        driver_path = ORACLE_WINDOWS_DRIVER_PATH
    elif platform.system() == "Linux":
        logging.log(
            logging.INFO,
            f"Linux OS detected, using driver path: {ORACLE_LINUX_DRIVER_PATH}",
            exc_info=True,
        )
        driver_path = ORACLE_LINUX_DRIVER_PATH
    else:
        logging.log(
            logging.ERROR,
            f"Unsupported OS: {platform.system()}",
            exc_info=True,
        )
        raise Exception(f"Unsupported OS: {platform.system()}")

    ezdriver_env_driver_path = os.getenv(ORACLE_DRIVER_ENV_VAR)

    if ezdriver_env_driver_path:
        logging.log(
            logging.INFO,
            f"Using driver path from environment variable: {ezdriver_env_driver_path}",
            exc_info=True,
        )
        driver_path = ezdriver_env_driver_path
    else:
        logging.log(
            logging.INFO,
            f"Using default driver path: {driver_path}",
            exc_info=True,
        )

    return driver_path


def extract_oracle_driver() -> bool:
    # If the oracle_driver.zip file is found then extract the driver
    logging.log(
        logging.INFO,
        f"Extracting Oracle Driver from: {ORACLE_DRIVER_ZIP}",
        exc_info=True,
    )
    driver_path = _get_oracle_driver_path_for_os()
    zip_path = join_path(get_current_dir(), ORACLE_DRIVER_ZIP)
    if check_file_exists(zip_path):
        logging.log(
            logging.INFO,
            f"Oracle Driver zip file found: {zip_path}",
            exc_info=True,
        )
        try:
            os.makedirs(driver_path, exist_ok=True)
            logging.log(
                logging.INFO,
                f"Oracle Driver directory created: {driver_path}",
                exc_info=True,
            )
        except Exception as e:
            logging.log(
                logging.ERROR,
                f"Error creating Oracle Driver directory: {e}",
                exc_info=True,
            )
            return False
        try:
            with ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(driver_path)
                logging.log(
                    logging.INFO,
                    f"Oracle Driver extracted to: {driver_path}",
                    exc_info=True,
                )
        except Exception as e:
            logging.log(
                logging.ERROR,
                f"Error extracting Oracle Driver: {e}",
                exc_info=True,
            )
            return False
        return True


def get_oracle_driver_path() -> str | None:
    logging.log(
        logging.INFO,
        f"Getting Oracle Driver Path",
        exc_info=True,
    )

    driver_path = join_path(_get_oracle_driver_path_for_os(), ORACLE_DRIVER_VERSION)

    logging.log(
        logging.INFO,
        f"Oracle Driver Path: {driver_path}",
        exc_info=True,
    )

    if check_file_exists(driver_path):
        logging.log(
            logging.INFO,
            f"Oracle Driver Path: {driver_path} exists",
            exc_info=True,
        )
        return driver_path
    else:
        extract_oracle_driver()
        if check_file_exists(driver_path):
            logging.log(
                logging.INFO,
                f"Oracle Driver Path: {driver_path} exists after extraction",
                exc_info=True,
            )
            return driver_path
        else:
            logging.log(
                logging.ERROR,
                f"Oracle Driver Path: {driver_path} does not exist after extraction",
                exc_info=True,
            )
            raise Exception(f"Oracle Driver Path: {driver_path} does not exist")
