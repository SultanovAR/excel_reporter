import numbers
import typing as tp
import math

import PIL
import xlsxwriter
import pandas as pd
from pandas.api.types import is_integer_dtype, is_float_dtype


from .utils import PIL2IOBytes


class BaseExcelReporter:
    CELL_WIDTH = 8.43  # 64 pixels - default Excel cell width
    PIXEL_CELL_WIDTH = 64
    PIXEL_ROW_HEIGHT = 20
    PIXEL_TO_WIDTH_RATIO = CELL_WIDTH / PIXEL_CELL_WIDTH
    PIXEL_CHAR_WIDTH = int(8 * 1.125)

    def __init__(self, excel_path,
                 pixel_sheet_width=1200,
                 max_cell_width=24,
                 min_cell_width=3.5):
        self.PIXEL_SHEET_WIDTH = pixel_sheet_width
        self.MAX_CELL_WIDTH = max_cell_width
        self.MIN_CELL_WIDTH = min_cell_width
        self.excel_path = excel_path
        self._wb = xlsxwriter.Workbook(excel_path)
        self._wb.nan_inf_to_errors = True
        self._sheet = None
        self._cursor = None

    def close(self):
        self._wb.close()

    def __del__(self):
        pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()

    def create_empty_sheet(self, name, default_cell_format={}):
        """Creates empty sheet and choose it as active sheet

        Parameters
        ----------
        name : name of sheet
        default_cell : dict of style cells. Read about Format in xlsx docs
        """
        self._sheet = self._wb.add_worksheet(name)
        cell_format = self._wb.add_format(default_cell_format)
        self._sheet.set_column('A:ZZ', self.CELL_WIDTH, cell_format)
        self.default_cell_format = self._wb.add_format(default_cell_format)
        self.set_active_sheet(name)

    def set_active_sheet(self, name, cursor=[0, 0]) -> None:
        self._sheet, self._cursor = self._wb.sheetnames[name], cursor

    def apply_shift(self, shift) -> None:
        self._cursor[0] += shift[0]
        self._cursor[1] += shift[1]

    def _text(self, string, format=None, shift=(2, 0)) -> None:
        text_pos = (self._cursor[0], self._cursor[1])
        cur_format = self._wb.add_format(format)
        self._sheet.write(*text_pos, string, cur_format)
        self.apply_shift(shift)

    def __printed_text_height(self, text) -> int:
        height = 0
        for line in text.split('\n'):
            height += (len(line) * self.PIXEL_CHAR_WIDTH //
                       self.PIXEL_SHEET_WIDTH)
            height += 1
        return height

    def __printed_value_width(self, value) -> float:
        return len(str(value)) * self.PIXEL_CHAR_WIDTH * self.PIXEL_TO_WIDTH_RATIO

    def __clip(self, value, lower, upper) -> numbers.Number:
        return lower if value < lower else upper if value > upper else value

    def _textbox(self, text, title=None,
                 header_format={}, textbox_format={}, shift=None) -> None:
        text_height = self.__printed_text_height(text)
        box_pos = self._cursor[0] + 1, self._cursor[1]
        self._sheet.insert_textbox(*box_pos, text,
                                   {'width': self.PIXEL_SHEET_WIDTH,
                                    'height': text_height * self.PIXEL_ROW_HEIGHT,
                                    'object_position': 1,
                                    **textbox_format})
        if title is not None:
            self._cursor_anchor = self._cursor.copy()
            self._text(title, header_format, shift=[0, 1])
            cell_header = len(str(title)) * \
                self.PIXEL_CHAR_WIDTH // self.PIXEL_CELL_WIDTH
            for _ in range(max(3, cell_header)):
                self._text('', header_format, shift=[0, 1])
            self._cursor = self._cursor_anchor
        if shift is None:
            shift = (text_height + 2, 0)
        self.apply_shift(shift)

    def _image(self, image: PIL.Image, image_name,
               max_pixel_size=(450, 800), image_options={},
               shift=None) -> None:
        scales = (lim / cur for cur, lim in zip(image.size, max_pixel_size))
        scale = min(1, *scales)
        im_pos = (self._cursor[0], self._cursor[1])
        self._sheet.insert_image(*im_pos, image_name,
                                 {'image_data': PIL2IOBytes(image),
                                  'x_scale': scale,
                                  'y_scale': scale,
                                  'object_position': 3,
                                  **image_options
                                  })
        if shift is None:
            shift = (math.ceil(image.size[1] *
                     scale / self.PIXEL_ROW_HEIGHT) + 1, 0)
        self.apply_shift(shift)

    # table

    def __get_col_width(self, col) -> int:
        if col in self._sheet.col_info:
            return self._sheet.col_info[col][0]
        return self.CELL_WIDTH

    def __adjust_col_width(self, col_ind, text, col_format) -> None:
        text_len = self.__printed_value_width(text)
        width = self.__clip(text_len, self.MIN_CELL_WIDTH, self.MAX_CELL_WIDTH)
        self._sheet.set_column(col_ind, col_ind, width, col_format)

    def __get_number_format(self, col_name, df):
        if isinstance(col_name, str) and ('%' in col_name):
            return {'num_format': '0.00%'}
        if is_integer_dtype(df[col_name]):
            return {'num_format': '0'}
        elif is_float_dtype(df[col_name]):
            return {'num_format': '0.000'}
        else:
            return None

    def __cast_to_str_except_numbers(self, value) -> tp.Union[str, numbers.Number]:
        if isinstance(value, numbers.Number):
            return value
        return str(value)

    def __get_table_item_format(self, cur_row, cur_col, df):
        num_format = self.__get_number_format(df.columns[cur_col], df)
        cur_format = self._wb.add_format(num_format)
        if (cur_row == len(df) - 1) and (cur_col == len(df.columns) - 1):
            cur_format.set_bottom()
            cur_format.set_right()
        elif (cur_row == len(df) - 1):
            cur_format.set_bottom()
        elif (cur_col == len(df.columns) - 1):
            cur_format.set_right()
        return cur_format

    def _table(self, df: pd.DataFrame,
               table_column_format,
               table_index_format,
               shift=None) -> None:
        col_format = self._wb.add_format(table_column_format)
        for j, col_name in enumerate(df.columns):
            cur_width = self.__get_col_width(j + self._cursor[1] + 1)
            need_width = self.__printed_value_width(col_name)
            if cur_width < need_width:
                self.__adjust_col_width(j + self._cursor[1] + 1,
                                        col_name, self.default_cell_format)
            self._sheet.write(self._cursor[0], j + self._cursor[1] + 1,
                              col_name, col_format)

        ind_format = self._wb.add_format(table_index_format)
        for i, (index, values) in enumerate(df.iterrows()):
            self._sheet.write(self._cursor[0] + i + 1, self._cursor[1],
                              index, ind_format)
            for j, val in enumerate(values):
                cur_format = self.__get_table_item_format(i, j, df)
                self._sheet.write(self._cursor[0] + i + 1,
                                  self._cursor[1] + j + 1,
                                  self.__cast_to_str_except_numbers(val),
                                  cur_format)
        if shift is None:
            shift = (len(df) + 2, 0)
        self.apply_shift(shift)
