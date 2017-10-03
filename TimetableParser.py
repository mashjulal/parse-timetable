import os
import re
import datetime
from pprint import pprint

import openpyxl

from TimetableSheet import TimetableSheet
from lessons.Lab import Lab
from lessons.Lecture import Lecture
from lessons.Practice import Practice


class TimetableParser:

    PATTERN_SPECIFIC_WEEKS = "([0-9,]+) н ([А-Яа-я \-A-Za-z]+)"
    PATTERN_EXCEPT_WEEKS = "кр ([0-9, ]+) н ([А-Яа-я \-A-Za-z]+)"

    FILE_PATH = os.getcwd() + "/generated_files/КБиСП 3 курс 1 сем.xlsx"

    def __init__(self, group_name):
        self.workbook = openpyxl.load_workbook(TimetableParser.FILE_PATH)
        self.timetable_sheet = self.workbook["Лист1"]
        self.group_name = group_name
        self.timetable = []

        self.create_timetable_list()
        self.get_timetable()

    def get_group_column(self):
        row = 2
        column = 6
        cnt = 0
        found = False

        cell_value = self.timetable_sheet.cell(row=row, column=column).value
        while cell_value:
            if self.group_name in cell_value:
                found = True
                break

            cnt += 1
            column += 4 if cnt % 3 != 0 else 10
            cell_value = self.timetable_sheet.cell(row=row, column=column).value

        return column if found else None

    def get_timetable(self):
        start_date = TimetableSheet.get_starting_date()
        end_date = TimetableSheet.get_end_date()
        week_count = ((end_date - start_date) // 7).days

        col = self.get_group_column()

        discipline_col = col
        type_col = col + 1
        lecturer_col = col + 2
        room_col = col + 3

        for row in range(4, 76):
            tp = str(self.timetable_sheet.cell(row=row, column=type_col).value).lower()
            lecturer = self.timetable_sheet.cell(row=row, column=lecturer_col).value
            room = self.timetable_sheet.cell(row=row, column=room_col).value
            weekday = (row - 4) // 12
            is_week_odd = (row - 4) % 2 == 0
            time = ((row - 4) % 12) // 2

            discipline = self.timetable_sheet.cell(row=row, column=discipline_col).value
            weeks = list(range(0, week_count))
            if discipline:
                if discipline.startswith("кр"):
                    discipline_match = \
                        re.fullmatch(TimetableParser.PATTERN_EXCEPT_WEEKS, discipline)
                    discipline = discipline_match.group(2)
                    weeks = [int(n) for n in range(1, 17)
                             if n not in discipline_match.group(1).split(",")]
                elif re.fullmatch(TimetableParser.PATTERN_SPECIFIC_WEEKS, discipline):
                    discipline_match = \
                        re.fullmatch(TimetableParser.PATTERN_SPECIFIC_WEEKS, discipline)
                    discipline = discipline_match.group(2)
                    weeks = [int(n) for n in discipline_match.group(1).split(",")]

            if tp == "прак" or tp == "пр":
                lesson = Practice(discipline, room, lecturer)
            elif tp == "лек":
                lesson = Lecture(discipline, room, lecturer)
            else:
                lesson = Lab(discipline, room, lecturer)

            if is_week_odd:
                weeks = [w for w in weeks if w % 2 == 0]
            else:
                weeks = [w for w in weeks if w % 2 == 1]

            for week in weeks:
                self.timetable[week-1][weekday][time] = lesson
        pprint(self.timetable)

    def create_timetable_list(self):
        start_date = TimetableSheet.get_starting_date()
        end_date = TimetableSheet.get_end_date()
        week_count = ((end_date - start_date) // 7).days

        self.timetable = [[[None for _ in range(6)]
                           for __ in range(6)]
                          for ___ in range(week_count+1)]
