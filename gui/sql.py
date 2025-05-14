import logging

import oracledb
from copy import deepcopy
from oracledb import Cursor, Connection

from features import utils


def get_connection(conn_string: str) -> Connection | None:
    logging.log(logging.INFO, "Getting Connection")

    if utils.get_oracle_driver_path():
        logging.log(logging.INFO, "Driver Path Found")
        try:
            oracledb.init_oracle_client(utils.get_oracle_driver_path())
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


def execute_query(conn: Connection, sql_query: str) -> list[list[str]] | None:
    logging.log(logging.INFO, f"Executing Query")
    logging.debug(f"Executing Query: {sql_query}")

    res = []

    if conn and sql_query.strip():
        cursor: Cursor = conn.cursor()

        try:
            cursor.execute(sql_query)
            col_names = [i[0] for i in cursor.description]

            data = deepcopy(cursor.fetchall())

            res.append([sql_query])
            res.append([])
            res.appedn(col_names)

            for row in data:
                res.append(row)

            return res
        except Exception as e:
            logging.log(logging.ERROR, f"Query Execution Failed: {e}")
            return None


def execute_bulk_query(
    conn: Connection, sql_queries: list[str]
) -> list[list[list[str]]] | None:
    logging.log(logging.INFO, f"Executing Bulk Query")
    logging.debug(f"Executing Bulk Query: {sql_queries}")

    res = []

    if conn and sql_queries:
        cursor: Cursor = conn.cursor()

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
