from typing import Tuple, List, TypeAlias
from enum import Enum
import io

class ACTION_TYPE(Enum):
    UNDO_LAST = "UNDO_LAST"
    CHANGE_SHEET = "CHANGE_SHEET"
    INSERT_IMAGE = "INSERT IMAGE"
    INSERT_QUERY_RES = "QUERY_RESULT"

PAYLOAD_TYPE: TypeAlias = None | str | Tuple[io.BytesIO, int] | List[List[List[str]]]

class Actions:

    def __init__(self, action_type: ACTION_TYPE, payload: PAYLOAD_TYPE):
        self.action_type: ACTION_TYPE = action_type
        self.payload = payload

    def getPayload(self):
        return self.payload


class InsertImage(Actions):

    def __init__(self, image_payload: Tuple[io.BytesIO, int]):
        super().__init__(ACTION_TYPE.INSERT_IMAGE, image_payload)

    
class ChangeSheet(Actions):

    def __init__(self, sheet_name: str):
        super().__init__(ACTION_TYPE.CHANGE_SHEET, sheet_name)

class InsertQueryResult(Actions):

    def __init__(self, sql_result: List[List[List[str]]]):
        super().__init__(ACTION_TYPE.INSERT_QUERY_RES, sql_result)
