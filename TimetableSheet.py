import xlwt
import datetime
import os


class TimetableSheet:

    FILE_PATH = os.getcwd() + "/generated_files/timetable.xls"

    WEEKDAYS = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]
    LECTURE_TIME = ["9-00", "10-30", "13-00", "14-40", "16-20", "18-00"]

    DISCIPLINE = "дисц"
    ROOM = "ауд"

    DATE_TEMPLATE = "%d.%m.%y"

    WEEKS_COUNT = 16

    def __init__(self):
        self.workbook = xlwt.Workbook()
        self.worksheet = self.workbook.add_sheet("Timetable")

    def generate_xls_file(self):
        self.generate_weekdays_rows()
        self.generate_lecture_times_column()
        self.generate_weeks()
        self.save()

    def generate_weekdays_rows(self):
        font = xlwt.Font()
        font.bold = True

        align = xlwt.Alignment()
        align.vert = xlwt.Alignment.VERT_CENTER
        align.horz = xlwt.Alignment.HORZ_CENTER
        align.rota = 90

        borders = xlwt.Borders()
        borders.bottom = xlwt.Borders.MEDIUM
        borders.left = xlwt.Borders.MEDIUM
        borders.right = xlwt.Borders.MEDIUM
        borders.top = xlwt.Borders.MEDIUM

        style = xlwt.XFStyle()
        style.font = font
        style.alignment = align
        style.borders = borders

        top_row = 2
        for weekday in TimetableSheet.WEEKDAYS:
            self.worksheet.write_merge(top_row, top_row+5, 0, 0, weekday, style=style)
            top_row += 7

    def generate_lecture_times_column(self):
        style = xlwt.XFStyle()

        font = xlwt.Font()
        font.bold = True

        borders = xlwt.Borders()
        borders.bottom = xlwt.Borders.MEDIUM
        borders.left = xlwt.Borders.MEDIUM
        borders.right = xlwt.Borders.MEDIUM
        borders.top = xlwt.Borders.MEDIUM

        align = xlwt.Alignment()
        align.horz = xlwt.Alignment.HORZ_CENTER

        style = xlwt.XFStyle()
        style.font = font
        style.borders = borders
        style.alignment = align

        top_row = 2
        for _ in range(len(TimetableSheet.WEEKDAYS)):
            for lt in TimetableSheet.LECTURE_TIME:
                self.worksheet.write(top_row, 1, lt, style=style)
                top_row += 1
            top_row += 1

    @staticmethod
    def get_starting_date():
        first_september = datetime.date(2017, 9, 1)
        first_september_weekday = first_september.weekday()
        starting_date = first_september - datetime.timedelta(first_september_weekday)
        return starting_date

    @staticmethod
    def get_end_date():
        first_september = datetime.date(2017, 9, 1)
        end_date = first_september + datetime.timedelta(7*16)
        end_date_weekday = end_date.weekday()
        end_date += datetime.timedelta(6 - end_date_weekday)
        return end_date

    def generate_weeks(self):
        font = xlwt.Font()
        font.bold = True

        align = xlwt.Alignment()
        align.horz = xlwt.Alignment.HORZ_CENTER

        borders = xlwt.Borders()
        borders.bottom = xlwt.Borders.MEDIUM
        borders.left = xlwt.Borders.MEDIUM
        borders.right = xlwt.Borders.MEDIUM
        borders.top = xlwt.Borders.MEDIUM

        style = xlwt.XFStyle()
        style.font = font
        style.alignment = align
        style.borders = borders


        left_column = 2

        start_date = TimetableSheet.get_starting_date()
        dt = TimetableSheet.get_starting_date()
        end_date = TimetableSheet.get_end_date()

        while dt < end_date:
            top_row = 0
            self.worksheet.write(top_row, left_column, TimetableSheet.DISCIPLINE, style=style)
            self.worksheet.write(top_row, left_column+1, TimetableSheet.ROOM, style=style)
            top_row += 1
            for _ in range(len(TimetableSheet.WEEKDAYS)):
                self.worksheet.write_merge(
                    top_row, top_row, left_column, left_column+1,
                    datetime.datetime.strftime(dt, TimetableSheet.DATE_TEMPLATE), style=style)
                top_row += 7
                dt += datetime.timedelta(1)
            week_number = ((dt - start_date) // 7).days + 1
            self.worksheet.write_merge(
                top_row, top_row, left_column, left_column+1, week_number, style=style)
            left_column += 2
            dt += datetime.timedelta(1)

    def save(self):
        self.workbook.save(TimetableSheet.FILE_PATH)