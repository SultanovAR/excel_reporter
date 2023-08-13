import typing as tp
import json

import xlsxwriter
import PIL
import pandas as pd

from .base_excel_reporter import BaseExcelReporter
from .utils import get_project_root


class ExcelReporter(BaseExcelReporter):
    NECESSARY_FORMAT_KEYS = [
        'default_cell',
        'sheet_title',
        'text_title',
        'textbox',
        'table_columns',
        'table_index']

    def __init__(self, excel_path,
                 theme_path=get_project_root() / 'themes/ctai_theme.json',
                 logo_path=get_project_root() / 'images/gpb_logo_white.png',
                 pixel_sheet_width=1200,
                 max_cell_width=24,
                 min_cell_width=3.5
                 ):
        super().__init__(excel_path, pixel_sheet_width, max_cell_width, min_cell_width)
        self._formats = self._read_theme(theme_path)
        self.logo_image = PIL.Image.open(logo_path)

    def _read_theme(self, theme_path) -> tp.List[xlsxwriter.workbook.Format]:
        formats = {}
        with open(theme_path, 'r') as f:
            formats = json.load(f)
            for key in self.NECESSARY_FORMAT_KEYS:
                if key not in formats:
                    raise ValueError(f"""Theme json-file should contain keys:{self.NECESSARY_FORMAT_KEYS}.
                                         define format for {key}""")
        return formats

    def _set_logo_title(self, title, sheet_name,
                        title_cell_height=2, logo_cell_width=2,
                        logo_offset={"x_offset": 8, "y_offset": 8},
                        title_prefix=' '):
        cur_format = self._wb.add_format(self._formats['sheet_title'])
        for i in range(title_cell_height):
            self._sheet.set_row(i, cell_format=cur_format)
        logo_h = self.PIXEL_ROW_HEIGHT * \
            title_cell_height - 2 * logo_offset['x_offset']
        logo_w = self.PIXEL_CELL_WIDTH * \
            logo_cell_width - 2 * logo_offset['y_offset']
        self._image(self.logo_image, f'logo_{sheet_name}',
                    max_pixel_size=(logo_w, logo_h),
                    image_options=logo_offset,
                    shift=(0, 0))
        self._cursor = [title_cell_height - 1, logo_cell_width]
        self._text(title_prefix + title, self._formats['sheet_title'],
                   shift=(0, 0))
        self._cursor = [title_cell_height, 1]

    def create_titled_sheet(self, sheet_name, title, description,
                            title_cell_height=2, logo_cell_width=2,
                            logo_offset={"x_offset": 8, "y_offset": 8},
                            title_prefix=' '):
        self.create_empty_sheet(sheet_name, self._formats['default_cell'])
        self._set_logo_title(title, sheet_name, title_cell_height,
                             logo_cell_width, logo_offset, title_prefix)
        self._textbox(description, None,
                      textbox_format=self._formats['textbox'])

    def insert_image(self, image: PIL.Image, image_name,
                     max_pixel_size=(450, 800), image_options={},
                     shift=None):
        self._image(image, image_name, max_pixel_size, image_options, shift)

    def insert_table(self, df: pd.DataFrame, shift=None):
        self._table(df, self._formats['table_columns'],
                    self._formats['table_index'], shift)

    def insert_textbox(self, text, title=None, shift=None) -> None:
        self._textbox(text, title, self._formats['text_title'],
                      self._formats['textbox'], shift)

    def insert_text(self, string, shift=(2, 0)):
        self._text(string, shift=shift)
