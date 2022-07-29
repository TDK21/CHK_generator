# This program generating cvs file for monthly time-workers
# Author TDK21
# ver 1.0

from pathlib import Path
from datetime import datetime
import pandas as pd
import calendar


def getFileName():
    date = datetime.date(datetime.now())
    split_date = str(date).split(sep="-")
    month = split_date[1]
    year = split_date[0]
    file_data = "check_" + month + "." + year
    return file_data


def getFullDirectory():
    f_name = getFileName()
    full_directory = Path('/', 'home', 'kuz', 'chk', f_name)
    return full_directory


def writeTable():
    days_list = ([])
    dinner = '00:45'
    i = 1
    row = 2
    count_rows = "="
    twd = "="
    split_date = str(datetime.date(datetime.now())).split(sep="-")
    month_numbers = (calendar.monthrange(int(split_date[0]), int(split_date[1])))  # get first day + number of days
    day = month_numbers[0]
    days_list.append(['DAY', 'DATE', 'IN', 'OUT', 'WORK', 'TOTAL', '', 'DINNER=', dinner])
    while i <= month_numbers[1]:
        if day > 4:
            days_list.append(['DAY', 'DATE', 'IN', 'OUT', 'WORK', 'TOTAL'])
            count_rows = "="
            row += 1
            day = 0
            i += 1
        else:
            if day == 0:
                days_list.append(["Monday", i, '', '', '=D' + str(row) + "-C" + str(row)])
                count_rows += "E" + str(row)
            if day == 1:
                days_list.append(["Tuesday", i, '', '', '=D' + str(row) + "-C" + str(row) + "-I1"])
                count_rows += "+E" + str(row)
            if day == 2:
                days_list.append(["Wednesday", i, '', '', '=D' + str(row) + "-C" + str(row) + "-I1"])
                count_rows += "+E" + str(row)
            if day == 3:
                days_list.append(["Thursday", i, '', '', '=D' + str(row) + "-C" + str(row)])
                count_rows += "+E" + str(row)
            if day == 4:
                days_list.append(["Friday", i, '', '', '=D' + str(row) + "-C" + str(row)])
                count_rows += "+E" + str(row)
                days_list.append(['', '', '', '', '', count_rows])
                row += 1
            twd += '+E' + str(row)
            day += 1
            row += 1
        i += 1
    days_list.append(['TOTAL WORKED HOURS IN MONTH', '', '', '', '', twd])
    return days_list


def main():
    current_directory = getFullDirectory()
    current_directory = str(current_directory) + ".xlsx"
    file = open(current_directory, "w")
    days_list = writeTable()
    df = pd.DataFrame(days_list)
    df.to_excel(current_directory, sheet_name='Timing', index=False, header=False)
    file.close()


if __name__ == "__main__":
    main()
