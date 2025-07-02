import re

from gui.ui_constants import MYSQL_URL_REGEX


def get_mysql_host(conn_string: str) -> str:
    match = re.search(MYSQL_URL_REGEX, conn_string)
    if match:
        return match.group(1)
    return ""


def get_mysql_user(conn_string: str) -> str:
    match = re.search(MYSQL_URL_REGEX, conn_string)
    if match:
        return match.group(2)
    return ""


def get_mysql_password(conn_string: str) -> str:
    match = re.search(MYSQL_URL_REGEX, conn_string)
    if match:
        return match.group(3)
    return ""


def get_mysql_database(conn_string: str) -> str:
    match = re.search(MYSQL_URL_REGEX, conn_string)
    if match:
        return match.group(4)
    return ""
