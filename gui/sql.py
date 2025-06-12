import logging
from copy import deepcopy
from typing import TypeAlias

import oracledb
import psycopg
import mysql.connector
from psycopg.connection import Connection as PgConnection
from oracledb import Connection as OracleConnection
from mysql.connector.pooling import PooledMySQLConnection
from mysql.connector.abstracts import MySQLConnectionAbstract

from gui.ui_constants import (
    POSTGRESQL_URL_VALIDATION_REGEX,
    ORACLE_URL_VALIDATION_REGEX,
    MYSQL_URL_VALIDATION_REGEX,
)
from gui.models import DatabaseEngine
import gui.utils as gui_utils
from features import utils


DatabaseConnection: TypeAlias = (
    PgConnection
    | OracleConnection
    | PooledMySQLConnection
    | MySQLConnectionAbstract
    | None
)


def get_oracle_connection(conn_string: str) -> DatabaseConnection:
    logging.log(logging.INFO, "Getting Connection")

    if utils.get_oracle_driver_path():
        logging.log(logging.INFO, "Driver Path Found")
        try:
            oracledb.init_oracle_client(utils.get_oracle_driver_path())  # type: ignore
        except Exception as e:
            logging.log(logging.ERROR, f"Connection Creation Failed: {e}")
    else:
        logging.log(logging.INFO, f"Cannot Find Oracle Driver Path")

    try:
        conn = oracledb.connect(conn_string)
        logging.log(logging.INFO, "Connection Created")
        return conn
    except Exception as e:
        logging.log(logging.ERROR, f"Connection Creation Failed: {e}")
        return None


def get_postgres_connection(conn_string: str) -> DatabaseConnection:
    logging.log(logging.INFO, "Getting PostgreSQL Connection")
    try:
        conn = psycopg.connect(conn_string)
        logging.log(logging.INFO, "PostgreSQL Connection Created")
        return conn
    except Exception as e:
        logging.log(logging.ERROR, f"PostgreSQL Connection Creation Failed: {e}")
        return None


def get_mysql_connection(conn_string: str) -> DatabaseConnection:
    logging.log(logging.INFO, "Getting MySQL Connection")
    try:
        conn = mysql.connector.connect(
            host=gui_utils.get_mysql_host(conn_string),
            user=gui_utils.get_mysql_user(conn_string),
            password=gui_utils.get_mysql_password(conn_string),
            database=gui_utils.get_mysql_database(conn_string),
            use_pure=False,
        )
        logging.log(logging.INFO, "MySQL Connection Created")
        return conn
    except Exception as e:
        logging.log(logging.ERROR, f"MySQL Connection Creation Failed: {e}")
        return None


def create_connection(engine: DatabaseEngine, conn_string: str) -> DatabaseConnection:
    logging.log(logging.INFO, f"Creating Connection for {engine.to_string()}")

    match engine:
        case DatabaseEngine.POSTGRESQL:
            logging.log(logging.DEBUG, f"PostgreSQL Connection String: {conn_string}")
            if not conn_string.strip() and not utils.validate_string_with_regex(
                conn_string, POSTGRESQL_URL_VALIDATION_REGEX
            ):
                logging.log(logging.ERROR, "Invalid PostgreSQL connection string")
                return None
            return get_postgres_connection(conn_string)
        case DatabaseEngine.ORACLE:
            logging.log(logging.DEBUG, f"Oracle Connection String: {conn_string}")
            if not conn_string.strip() and not utils.validate_string_with_regex(
                conn_string, ORACLE_URL_VALIDATION_REGEX
            ):
                logging.log(logging.ERROR, "Invalid Oracle connection string")
                return None
            return get_oracle_connection(conn_string)
        case DatabaseEngine.MYSQL:
            logging.log(logging.DEBUG, f"MySQL Connection String: {conn_string}")
            if not conn_string.strip() and not utils.validate_string_with_regex(
                conn_string, MYSQL_URL_VALIDATION_REGEX
            ):
                logging.log(logging.ERROR, "Invalid MySQL connection string")
                return None
            return get_mysql_connection(conn_string)


def execute_query(conn: DatabaseConnection, sql_query: str) -> list[list[str]] | None:
    logging.log(logging.INFO, f"Executing Query")
    logging.debug(f"Executing Query: {sql_query}")

    res: list[list[str]] = []

    if conn and sql_query.strip():
        cursor = conn.cursor()

        try:
            cursor.execute(sql_query)  # type: ignore
            col_names: list[str] = [row[0] for row in cursor.description]  # type: ignore

            data: list[list[str]] = deepcopy(cursor.fetchall())  # type: ignore

            res.append([sql_query])
            res.append([])
            res.append(col_names)

            for row in data:
                res.append(list(row))

            return res
        except Exception as e:
            logging.log(logging.ERROR, f"Query Execution Failed: {e}")
            return None


def execute_bulk_query(
    conn: DatabaseConnection, sql_queries: list[str]
) -> list[list[list[str]]] | None:
    logging.log(logging.INFO, f"Executing Bulk Query")
    logging.debug(f"Executing Bulk Query: {sql_queries}")

    res: list[list[list[str]]] = []

    if conn and sql_queries:

        for query in sql_queries:
            exec_res = execute_query(conn, query)
            if exec_res:
                res.append(exec_res)
            else:
                logging.log(logging.INFO, f"Invalid query on: {query}")

    if res:
        logging.log(logging.INFO, f"Query Executed Successfully")
        return res

    logging.log(logging.ERROR, f"Query Execution Failed")
    return None
