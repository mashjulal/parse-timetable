import datetime


def get_starting_date():
    first_september = datetime.date(2017, 9, 1)
    first_september_weekday = first_september.weekday()
    starting_date = first_september - datetime.timedelta(first_september_weekday)
    return starting_date

def get_end_date():
    first_september = datetime.date(2017, 9, 1)
    end_date = first_september + datetime.timedelta(7 * 16)
    end_date_weekday = end_date.weekday()
    end_date += datetime.timedelta(6 - end_date_weekday)
    return end_date


def get_week_count():
    start_date = get_starting_date()
    end_date = get_end_date()
    week_count = ((end_date - start_date) // 7).days
    return week_count


def get_acronym(words):
    return "".join([word[0].upper() for word in words.split()]) if words else ""
