from enum import Enum


class DatabaseEngine(Enum):
    POSTGRESQL = "postgresql"
    MYSQL = "mysql"
    ORACLE = "oracle_db"

    def to_string(self) -> str:
        return self.value

    @staticmethod
    def from_string(engine_str: str) -> "DatabaseEngine":
        engine_str = engine_str.lower()
        if engine_str == "postgresql":
            return DatabaseEngine.POSTGRESQL
        elif engine_str == "mysql":
            return DatabaseEngine.MYSQL
        elif engine_str == "oracle_db":
            return DatabaseEngine.ORACLE
        else:
            raise ValueError(f"Unknown database engine: {engine_str}")

    @staticmethod
    def get_supported_engines() -> list[str]:
        return [engine.value for engine in DatabaseEngine]


class DatabaseConnection:
    def __init__(self, engine: DatabaseEngine, conn_string: str):
        self.engine = engine
        self.conn_string = conn_string

    def __repr__(self):
        return (
            f"DatabaseConnection(engine={self.engine}, conn_string={self.conn_string})"
        )

    def __str__(self):
        return f"Database Connection: {self.engine.value} - {self.conn_string}"
