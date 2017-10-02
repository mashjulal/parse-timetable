import xlwt
import datetime
import os


class TimetableSheet:

    FILE_PATH = os.getcwd() + "/generated_files/timetable.xls"

    WEEKDAYS = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]
    LECTURE_TIME = ["9-00", "10-30", "13-00", "14-40", "16-20", "18-00"]

    def __init__(self):
        self.workbook = xlwt.Workbook()
        self.worksheet = self.workbook.add_sheet("Timetable")

    def generate_xls_file(self):
        start_date = TimetableSheet.get_starting_date()
        self.generate_weekdays_rows()
        self.generate_lecture_times_column()
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

    def save(self):
        self.workbook.save(TimetableSheet.FILE_PATH)