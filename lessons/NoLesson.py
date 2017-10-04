import xlwt

from lessons.Lesson import Lesson


class NoLesson(Lesson):

    INSTANCE = None

    def __init__(self):
        super().__init__(None, None, None)

    @staticmethod
    def get_instance():
        if not NoLesson.INSTANCE:
            NoLesson.INSTANCE = NoLesson()
        return NoLesson.INSTANCE

    def get_cell_style(self):
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

        return style