import xlwt

from Styles import Styles
from lessons.Lesson import Lesson


class NoLesson(Lesson):

    INSTANCE = None
    STYLE = Styles.TITLE

    def __init__(self):
        super().__init__(None, None, None)

    @staticmethod
    def get_instance():
        if not NoLesson.INSTANCE:
            NoLesson.INSTANCE = NoLesson()
        return NoLesson.INSTANCE
