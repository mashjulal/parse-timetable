import xlwt
import datetime
import os

from Styles import Styles
from TimetableParser import TimetableParser
import utils


class TimetableSheet:

    FILE_PATH = os.getcwd() + "/generated_files/timetable.xls"

    WEEKDAYS = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]
    LECTURE_TIME = ["9-00", "10-30", "13-00", "14-40", "16-20", "18-00"]

    DISCIPLINE = "дисц"
    ROOM = "ауд"

    DATE_TEMPLATE = "%d.%m.%y"

    WEEKS_COUNT = utils.get_week_count()

    def __init__(self):
        self.workbook = xlwt.Workbook()
        self.worksheet = self.workbook.add_sheet("Timetable")
        self.timetable_parser = TimetableParser("БНБО-01-15")

    def generate_xls_file(self):
        self.generate_weekdays_rows()
        self.generate_lecture_times_column()
        self.generate_weeks()

        self.timetable_parser.parse()
        self.generate_timetable()

        self.save()

    def generate_weekdays_rows(self):
        top_row = 2
        for weekday in TimetableSheet.WEEKDAYS:
            self.worksheet.write_merge(top_row, top_row + 5, 0, 0,
                                       weekday,
                                       style=Styles.WEEKDAY)
            top_row += 7

    def generate_lecture_times_column(self):
        top_row = 2
        for _ in range(len(TimetableSheet.WEEKDAYS)):
            for lt in TimetableSheet.LECTURE_TIME:
                self.worksheet.write(top_row, 1, lt, style=Styles.TITLE)
                top_row += 1
            top_row += 1

    def generate_weeks(self):
        left_column = 2

        first_non_academic_day = utils.get_first_non_academic_date()
        dt = first_non_academic_day
        last_non_academic_day = utils.get_last_non_academic_day()

        while dt < last_non_academic_day:
            top_row = 0
            self.worksheet.write(top_row, left_column,
                                 TimetableSheet.DISCIPLINE,
                                 style=Styles.TITLE)
            self.worksheet.write(top_row, left_column+1,
                                 TimetableSheet.ROOM,
                                 style=Styles.TITLE)
            top_row += 1
            for _ in range(len(TimetableSheet.WEEKDAYS)):
                self.worksheet.write_merge(
                    top_row, top_row, left_column, left_column+1,
                    datetime.datetime.strftime(dt, TimetableSheet.DATE_TEMPLATE),
                    style=Styles.TITLE)
                top_row += 7
                dt += datetime.timedelta(1)
            week_number = (dt - first_non_academic_day).days // 7 + 1
            self.worksheet.write_merge(
                top_row, top_row, left_column, left_column+1,
                week_number,
                style=Styles.TITLE)
            left_column += 2
            dt += datetime.timedelta(1)

    def generate_timetable(self):
        column = 2
        for week_index, week in enumerate(self.timetable_parser.timetable):
            row = 2
            for day_index, day in enumerate(week):
                for lesson_index, lesson in enumerate(day):
                    self.write_into_cell(row, column, lesson)
                    row += 1
                row += 1
            column += 2

    def write_into_cell(self, row, column, lesson):
        self.worksheet.write(
            row, column, lesson.discipline, style=lesson.STYLE)
        self.worksheet.write(
            row, column + 1, lesson.room, style=lesson.STYLE)

    def save(self):
        self.workbook.save(TimetableSheet.FILE_PATH)
