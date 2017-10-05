from Styles import Styles
from lessons.Lesson import Lesson


class NonAcademic(Lesson):

    STYLE = Styles.NON_ACADEMIC
    INSTANCE = None

    def __init__(self):
        super().__init__(None, None, None)

    @staticmethod
    def get_instance():
        if not NonAcademic.INSTANCE:
            NonAcademic.INSTANCE = NonAcademic()
        return NonAcademic.INSTANCE
