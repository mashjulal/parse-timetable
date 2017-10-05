import utils
from Styles import Styles


class AbstractLesson:

    STYLE = None

    def __init__(self, discipline, room, lecturer):
        self.discipline = utils.get_acronym(discipline)
        self.room = room
        self.lecturer = lecturer


class NoLesson(AbstractLesson):

    INSTANCE = None
    STYLE = Styles.TITLE

    def __init__(self):
        super().__init__(None, None, None)

    @staticmethod
    def get_instance():
        if not NoLesson.INSTANCE:
            NoLesson.INSTANCE = NoLesson()
        return NoLesson.INSTANCE


class NonAcademic(AbstractLesson):

    STYLE = Styles.NON_ACADEMIC
    INSTANCE = None

    def __init__(self):
        super().__init__(None, None, None)

    @staticmethod
    def get_instance():
        if not NonAcademic.INSTANCE:
            NonAcademic.INSTANCE = NonAcademic()
        return NonAcademic.INSTANCE


class Practice(AbstractLesson):

    STYLE = Styles.PRACTICE


class Lecture(AbstractLesson):

    STYLE = Styles.LECTURE


class Lab(AbstractLesson):

    STYLE = Styles.LAB
