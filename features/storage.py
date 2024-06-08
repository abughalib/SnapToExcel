from enum import Enum
from typing import Deque
from collections import deque
from .actions import ACTION_TYPE, Actions


class Storage:

    def __init__(self) -> None:
        self.stack: Deque[Actions] = deque()

    def _undo(self) -> str:
        if self.stack:
            action = self.stack.pop()
            action_type: str = ""

            if action.action_type == ACTION_TYPE.CHANGE_SHEET:
                action_type = "Sheet Name Removed"
            elif action.action_type == ACTION_TYPE.INSERT_IMAGE:
                action_type = "Last Screenshot Removed"
            elif action.action_type == ACTION_TYPE.INSERT_QUERY_RES:
                action_type = "Last SQL Query Result Removed"

            return action_type

        return "No Action Taken Yet"

    def _add(self, action: Actions) -> str:
        self.stack.append(action)
        return "Action Added"

    def get_stack(self):
        return self.stack

    def get_last(self) -> Actions | None:
        if self.stack:
            return self.stack[-1]

        return None

    def execute(self, action: Actions) -> str:
        if action.action_type == ACTION_TYPE.UNDO_LAST:
            return self._undo()
        elif action.action_type in [
            ACTION_TYPE.CHANGE_SHEET,
            ACTION_TYPE.INSERT_IMAGE,
            ACTION_TYPE.INSERT_QUERY_RES,
        ]:
            if action:
                return self._add(action)
            return "Insufficient Args"

        return "Unknown Error"
