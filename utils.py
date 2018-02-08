import datetime

TOTAL_ACADEMIC_DAYS = 7 * 16


def get_first_academic_day():
    cur_date = datetime.datetime.now()
    cur_year, cur_month = cur_date.year, cur_date.month

    if cur_month in range(8, 12 + 1):
        first_academic_day = datetime.date(cur_year, 9, 1)
    else:
        date = datetime.date(cur_year, 2, 1)
        date += datetime.timedelta(8)
        first_academic_day = date

    return first_academic_day


def get_last_academic_day():
    first_academic_day = get_first_academic_day()
    return first_academic_day + datetime.timedelta(7*16)


def get_first_non_academic_date():
    first_academic_day = get_first_academic_day()
    first_non_academic_day = \
        first_academic_day - datetime.timedelta(first_academic_day.weekday())
    return first_non_academic_day


def get_last_non_academic_day():
    first_academic_day = get_first_academic_day()
    last_academic_day = first_academic_day + datetime.timedelta(TOTAL_ACADEMIC_DAYS)
    last_non_academic_day = \
        last_academic_day - datetime.timedelta(6 - last_academic_day.weekday())
    return last_non_academic_day


def get_week_count():
    first_non_academic_day = get_first_non_academic_date()
    last_non_academic_day = get_last_non_academic_day()
    week_count = (last_non_academic_day - first_non_academic_day).days // 7
    return week_count


def get_acronym(words):
    if not words:
        return ""
    elif all(c.isupper() for c in words):
        return words
    return "".join([word[0].upper() for word in words.split()])

