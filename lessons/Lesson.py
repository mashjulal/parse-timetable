import utils


class Lesson:

    STYLE = None

    def __init__(self, discipline, room, lecturer):
        self.discipline = utils.get_acronym(discipline)
        self.room = room
        self.lecturer = lecturer