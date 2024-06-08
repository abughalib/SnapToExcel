import io
import logging
import pandas as pd
from PIL import Image as PILImage
from .storage import Storage
from .actions import ACTION_TYPE, Actions, InsertImage, InsertQueryResult
from typing import Tuple, cast
from .utils import join_path
from typing import List


DEFAULT_SCREENSHOT_HEIGHT = 60
PIXEL_PER_ROW = 20


class XlxsSheet:

    def __init__(
        self,
        folder_path: str,
        file_name: str,
        start_row: int,
        start_column: str,
        sheet_name="Sheet1",
        row_seperation: int = DEFAULT_SCREENSHOT_HEIGHT,
    ) -> None:
        logging.log(logging.INFO, f"XlxsSheet Constructor")
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.start_row = start_row
        self.next_row = start_row
        self.start_column = start_column
        self.row_seperation = row_seperation
        self.saved = False
        self.file_path = join_path(folder_path, self.file_name)
        self.storage = Storage()
        self.sheet = pd.DataFrame()
        self.sheet_index = dict()
        self.sheet_index[self.sheet_name] = start_row
        self.writer = pd.ExcelWriter(self.file_path, engine="xlsxwriter")

        logging.log(
            logging.INFO,
            f"""XlxsSheet
                file_name: {self.file_name}
                sheet_name: {self.sheet_name}
                next_row: {self.next_row}
                file_path: {self.file_path}
            """,
        )

    def _add_sheet(self, sheet_name: str) -> None:
        """Add Sheet to excel file"""
        logging.log(logging.INFO, f"Adding Sheet: {sheet_name}")
        self.sheet_name = sheet_name
        self.next_row = self.next_row
        self.writer.book.add_worksheet(sheet_name)  # type: ignore

    def add_image_to_queue(self, image: PILImage.Image) -> None:
        # Save the image to an in-memory buffer
        logging.log(logging.INFO, "Saving the image to memory")
        image_buffer = io.BytesIO()
        image.save(image_buffer, format="PNG")
        logging.log(logging.INFO, "Saved the image to memory")

        self.storage.execute(InsertImage((image_buffer, image.height)))

    def undo_last_action(self):
        logging.log(logging.INFO, "Removing last image")
        res = self.storage.execute(Actions(ACTION_TYPE.UNDO_LAST, None))
        return res

    def add_image(self, image: Tuple[io.BytesIO, int]) -> None:
        """Adds an image to a new row in the worksheet"""
        image_buffer, image_height = image
        logging.log(logging.INFO, f"Adding Image to sheet: {self.sheet_name}")
        if self.sheet_name not in self.writer.sheets:
            logging.log(
                logging.INFO, f"Sheet Name {self.sheet_name} not found in workbook"
            )
            self._add_sheet(self.sheet_name)
        logging.log(logging.INFO, f"Adding Image: {self.sheet_name}")
        worksheet = self.writer.sheets[self.sheet_name]

        cell_range = f"{self.start_column}{self.next_row}"
        logging.log(logging.INFO, f"Next Image is Inserting to: {cell_range}")

        worksheet.insert_image(
            cell_range,
            "",
            options={"image_data": image_buffer, "image_data_id": self.sheet_name},
        )
        logging.log(logging.INFO, "Worksheet image inserted")
        self.next_row += self.row_seperation + self._screenshot_row_height(image_height)
        self.sheet_index[self.sheet_name] = self.next_row

    def insert_list_values_queue(self, values: List[List[List[str]]]) -> None:
        self.storage.execute(InsertQueryResult(values))

    def _insert_list_values(self, values: List[List[List[str]]]) -> None:
        if self.sheet_name not in self.writer.sheets:
            logging.log(
                logging.INFO, f"Sheet Name: {self.sheet_name} not found in workbook"
            )
            self._add_sheet(self.sheet_name)

        logging.log(logging.INFO, f"Adding Image: {self.sheet_name}")
        worksheet = self.writer.sheets[self.sheet_name]

        for rows in values:
            for row in rows:
                to_write = []
                for data in row:
                    try:
                        data_str = str(data)
                        to_write.append(data_str)
                    except AttributeError:
                        logging.log(logging.ERROR, "Cannot result to string")

                    cell_range = f"{self.start_column}{self.next_row}"

                    try:
                        worksheet.write_row(cell_range, to_write)

                    except Exception as e:
                        logging.log(logging.INFO, f"Failed to add data: {to_write}")

                    self.next_row += 1
                self.next_row += self.row_seperation

    def _screenshot_row_height(self, height: int) -> int:
        return int(height // PIXEL_PER_ROW)

    def change_sheet(self, new_name: str) -> None:
        logging.log(
            logging.INFO, f"Adding Image first then change sheet to: {self.sheet_name}"
        )
        self.storage.execute(Actions(ACTION_TYPE.CHANGE_SHEET, new_name))
        logging.log(logging.INFO, "Changed and added new sheet")

    def change_row_seperation(
        self, seperation: int = DEFAULT_SCREENSHOT_HEIGHT
    ) -> None:
        self.row_seperation = seperation

    def change_position(self, row_name: int, col_name: str) -> None:
        self.start_row = row_name
        self.start_column = col_name

    def save_sheet(self) -> bool:
        logging.log(logging.INFO, "Saving Excel Sheet")
        stack = self.storage.get_stack()

        for action in stack:
            if action.action_type == ACTION_TYPE.INSERT_IMAGE:
                self.add_image(cast(Tuple[io.BytesIO, int], action.payload))
            elif action.action_type == ACTION_TYPE.INSERT_QUERY_RES:
                self._insert_list_values(
                    cast(List[List[List[str]]], action.getPayload())
                )
            elif action.action_type == ACTION_TYPE.CHANGE_SHEET:
                self.sheet_name = cast(str, action.getPayload())

        try:
            self.sheet.to_excel(excel_writer=self.writer, sheet_name=self.sheet_name)
            self.writer.close()
            self.saved = True
            logging.log(
                logging.INFO, f"Excel File: {self.file_name} Saved at: {self.file_path}"
            )
            return True
        except Exception as e:
            logging.log(logging.INFO, f"Error in saving file: {e}")
            return False
