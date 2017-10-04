import datetime
import os
import re

import openpyxl

import utils
from lessons.Lab import Lab
from lessons.Lecture import Lecture
from lessons.NoLesson import NoLesson
from lessons.Practice import Practice


class TimetableParser:

    PATTERN_SPECIFIC_WEEKS = "([0-9, ]+) н ([А-Яа-я \-A-Za-z]+)"
    PATTERN_EXCEPT_WEEKS = "кр ([0-9, ]+) н ([А-Яа-я \-A-Za-z]+)"

    FILE_PATH = os.getcwd() + "/generated_files/official_timetable.xlsx"

    def __init__(self, group_name):
        workbook = openpyxl.load_workbook(TimetableParser.FILE_PATH)
        self.timetable_sheet = workbook["Лист1"]
        self.group_name = group_name
        self.timetable = []

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
        self.init_timetable()
        week_count = utils.get_week_count()

        first_study_day = datetime.date(2017, 9, 1)
        last_study_day = first_study_day + datetime.timedelta(7*16)
        first_day_pos = first_study_day.weekday()
        last_day_pos = week_count * 6 + last_study_day.weekday()

        col = self.get_group_column()

        discipline_col = col
        type_col = col + 1
        lecturer_col = col + 2
        room_col = col + 3

        for row in range(4, 76):
            discipline = self.timetable_sheet.cell(row=row, column=discipline_col).value

            if not discipline:
                continue

            tp = self.timetable_sheet.cell(row=row, column=type_col).value
            lecturer = self.timetable_sheet.cell(row=row, column=lecturer_col).value
            room = self.timetable_sheet.cell(row=row, column=room_col).value
            weekday = (row - 4) // 12
            is_week_odd = (row - 4) % 2 == 0
            time = ((row - 4) % 12) // 2

            weeks = filter(
                lambda week_i: first_day_pos <= (week_i - 1) * 6 + weekday < last_day_pos,
                list(range(1 if is_week_odd else 2, week_count + 2, 2)))

            if discipline.startswith("кр"):
                discipline_match = \
                    re.fullmatch(TimetableParser.PATTERN_EXCEPT_WEEKS, discipline)
                discipline = discipline_match.group(2)
                weeks = filter(
                    lambda week_i: week_i not in discipline_match.group(1).split(","),
                    weeks)
            elif re.fullmatch(TimetableParser.PATTERN_SPECIFIC_WEEKS, discipline):
                discipline_match = \
                    re.fullmatch(TimetableParser.PATTERN_SPECIFIC_WEEKS, discipline)
                discipline = discipline_match.group(2)
                weeks = [int(n) for n in discipline_match.group(1).split(",")]

            lesson = TimetableParser.get_lesson_by_type(tp, discipline, room, lecturer)

            for week in weeks:
                self.timetable[week-1][weekday][time] = lesson

    def init_timetable(self):
        week_count = utils.get_week_count()

        self.timetable = [[[NoLesson.get_instance() for _ in range(6)]
                           for _ in range(6)]
                          for _ in range(week_count+1)]

    @staticmethod
    def get_lesson_by_type(tp, discipline, room, lecturer):
        lesson = None
        if re.fullmatch("(практ|пр)", tp, flags=re.I):
            lesson = Practice(discipline, room, lecturer)
        elif re.fullmatch("лек", tp, flags=re.I):
            lesson = Lecture(discipline, room, lecturer)
        elif re.fullmatch("лаб", tp, flags=re.I):
            lesson = Lab(discipline, room, lecturer)

        return lesson
