import os
from assignment import *
from families_data import all_families


def main():
    start = input("Enter the start date(DD/MM/YYYY): ")
    end = input("Enter the end date(DD/MM/YYYY): ")
    start_day = int(start.split('/')[0])
    start_month = int(start.split('/')[1])
    start_year = int(start.split('/')[2])
    end_day = int(end.split('/')[0])
    end_month = int(end.split('/')[1])
    end_year = int(end.split('/')[2])
    dates = get_dates(start_year, start_month,
                      start_day, end_year, end_month, end_day)
    timetable = leaders_assignment(dates, all_families)
    save_timetable_to_excel(timetable)
    if os.path.exists('kamau_family_contribution'):
        rewrite_timetable_in_excel(timetable)
    else:
        save_timetable_to_excel(timetable)


if __name__ == "__main__":
    main()
