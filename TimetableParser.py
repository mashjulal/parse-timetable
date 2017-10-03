import openpyxl
import os

from Lesson import Lesson
from TimetableSheet import TimetableSheet


class TimetableParser:

    FILE_PATH = os.getcwd() + "/generated_files/КБиСП 3 курс 1 сем.xlsx"

    def __init__(self, group_name):
        self.workbook = openpyxl.load_workbook(TimetableParser.FILE_PATH)
        self.timetable_sheet = self.workbook["Лист1"]
        self.group_name = group_name

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
        lessons = []
        col = self.get_group_column()

        discipline_col = col
        type_col = col + 1
        lecturer_col = col + 2
        room_col = col + 3

        for row in range(4, 76):
            discipline = self.timetable_sheet.cell(row=row, column=discipline_col).value
            tp = self.timetable_sheet.cell(row=row, column=type_col).value
            lecturer = self.timetable_sheet.cell(row=row, column=lecturer_col).value
            room = self.timetable_sheet.cell(row=row, column=room_col).value
            weekday = TimetableSheet.WEEKDAYS[(row - 4) // 12]
            is_week_odd = (row - 4) % 2 == 0
            time = TimetableSheet.LECTURE_TIME[((row - 4) % 12) // 2]

            lesson = Lesson(discipline, weekday, time, room, tp, is_week_odd, lecturer)
            lessons.append(lesson)

        for lesson in lessons:
            print(lesson.discipline, lesson.type, lesson.weekday, lesson.time, lesson.room, lesson.is_week_odd, lesson.lecturer)





