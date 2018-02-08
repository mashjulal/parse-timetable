from TimetableSheet import TimetableSheet

if __name__ == "__main__":
    # group_name = input("Input your group name: ")
    group_name = "БНБО-01-15"

    timetable_sheet = TimetableSheet()
    timetable_sheet.generate_timetable_for_group(group_name)