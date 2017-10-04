import xlwt


class Practice:

    def __init__(self, discipline, room, lecturer):
        self.discipline = discipline
        self.room = room
        self.lecturer = lecturer

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

        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN

        style = xlwt.easyxf("pattern: pattern solid, fore_color green")
        style.font = font
        style.alignment = align
        style.borders = borders

        return style