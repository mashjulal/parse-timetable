import openpyxl
from openpyxl.utils import get_column_letter
import os


class TimetableParser:

    FILE_PATH = os.getcwd() + "/generated_files/КБиСП 3 курс 1 сем.xlsx"

    def __init__(self):
        self.workbook = openpyxl.load_workbook(TimetableParser.FILE_PATH)
        self.timetable_sheet = self.workbook["Лист1"]

    def get_group_column(self, group_name):
        row = 2
        column = 6
        cnt = 0
        found = False

        cell_value = self.timetable_sheet.cell(row=row, column=column).value
        while cell_value:
            if group_name in cell_value:
                found = True
                break

            cnt += 1
            column += 4 if cnt % 3 != 0 else 10
            cell_value = self.timetable_sheet.cell(row=row, column=column).value

        return column if found else None


