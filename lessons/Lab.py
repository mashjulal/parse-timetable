import xlwt

from lessons.Lesson import Lesson


class Lab(Lesson):

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

        style = xlwt.easyxf("pattern: pattern solid, fore_color blue")
        style.font = font
        style.alignment = align
        style.borders = borders

        return style