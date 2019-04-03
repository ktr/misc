"""
We receive a baseball schedule from little league coach in Excel format. That
is not useful for me since I rely on my calendar to remind me of events. The
script below takes that spreadsheet and turns it into a CSV file that can be
uploaded into Outlook. There are 2 separate entries made for each game: 1) the
actual game itself, and 2) travel time to the game (with a 5min reminder).
"""

import csv
import datetime
import sys
import sxl

team_name = 'TBD'
game_duration = datetime.timedelta(hours=1, minutes=30)
reminder = datetime.timedelta(minutes=5)
travel_time = datetime.timedelta(minutes=30)
subj_fmt = 'Little league game: {} @ {} ({})'

src = r'H:\lib\python\misc\outlook_calendar_upload.py'

wb = sxl.Workbook(r'H:\2019 AA Schedule v1 2.18.19.xlsx')
ws = wb.sheets['2019 AA Schedule']

columns = (
    'Subject',          # reqd
    'Location',         # not reqd
    'Start Time',       # not reqd
    'Start Date ',      # reqd
    'End Time',         # not reqd
    'End Date ',        # reqd
    'All Day Event',    # not reqd; yes/no
    'Reminder On/Off',  # not reqd; yes/no
    'Reminder Date ',   # not reqd
    'Reminder Time',    # not reqd
    'Categories',       # not reqd
    'Description',      # not reqd
    'Private',          # not reqd; yes/no
    'Sensitivity',      # not reqd
    'Show Time As',     # not reqd; 1- Tentative/2- Busy/3- Free/4- Out of Office
)

header = ws.rows[3][0]
header_m = { name : pos for pos, name in enumerate(header) }
get = lambda col, row: row[header_m[col]]
with open('H:/baseball_calendar.csv', 'w', newline='') as csv_file:
    writer = csv.writer(csv_file)
    writer.writerow(columns)
    for i, row in enumerate(ws.rows[4:]):
        away = get('AWAY', row)
        home = get('HOME', row)
        if team_name not in (away, home):
            continue
        notes = get('Notes', row).strip()
        date = get('DATE', row).date()
        subj = subj_fmt.format(away, home, notes)
        # get game start time
        time_s = get('TIME', row)
        # turn start time into a date (so we can add time to it)
        time_d = datetime.datetime(1, 1, 1, time_s.hour, time_s.minute)
        # get game end time by adding game duration to start time (chop off
        # date component)
        time_e = (time_d + game_duration).time()
        # also useful for travel time
        time_ts = (time_d - travel_time).time()
        time_te = time_s
        time_tr = (time_d - travel_time - reminder).time()

        # Game
        writer.writerow([
            subj,                       # subj
            get('FIELD', row).strip(),  # location
            time_s,                     # start_time
            date,                       # start_date
            time_e,                     # end_time
            date,                       # end_date
            'No',                       # all_day
            'No',                       # reminder (will add to travel time)
            '',                         # reminder_date
            '',                         # reminder_time
            'Home',                     # categories
            src,                        # description
            '',                         # private
            '',                         # sensitivity
            '4- Out of Office',         # show_time_as
        ])
        # Travel time
        writer.writerow([
            'Travel time - baseball',   # subj
            '',                         # location
            time_ts,                    # start_time
            date,                       # start_date
            time_te,                    # end_time
            date,                       # end_date
            'No',                       # all_day
            'Yes',                      # reminder
            date,                       # reminder_date
            time_tr,                    # reminder_time
            'Home',                     # categories
            src,                        # description
            '',                         # private
            '',                         # sensitivity
            '4- Out of Office',         # show_time_as
        ])
