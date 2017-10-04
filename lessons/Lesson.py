import utils


class Lesson:

    def __init__(self, discipline, room, lecturer):
        self.discipline = utils.get_acronym(discipline)
        self.room = room
        self.lecturer = lecturer

    def get_cell_style(self):
        pass