import os
import webbrowser
from crontab import CronTab
import xlrd
from datetime import date, datetime, timedelta, timezone
import clipboard

########################## OPEN LINK #############################

file_name = "/Users/peter/Desktop/Programming/Python/autoZoom/schedule.xlsx"
schedule = xlrd.open_workbook(file_name)
sheet = schedule.sheet_by_index(0)
print(" ")


def get_weekday_time() -> list:
    """
    Returns the weekday and time in a list in the format [weekday, time]
    Weekday is in the list days = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]

    >>> get_weekday_time()
    ["MON", "12:01pm"]
    >>> get_weekday_time()
    ["WED", "11:53am"]
    """
    days = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
    return [days[date.today().weekday()], check_time()]

def check_time():
    """
    Returns current UTC time
    """
    current_utc_time = datetime.now().strftime("%H:%M")
    # print("Checking Current UTC Time: " + current_utc_time)
    return current_utc_time

def convert_utc_time(time_to_change):
    """
    Returns time_to_change in UTC time
    """
    converted_time = (datetime.strptime(
        time_to_change, "%H:%M") - timedelta(hours=8)).strftime("%H:%M")
    # print("Checking time (converted): " + converted_time)
    return converted_time


def get_zoom_details(current_time, current_day):
    """
    Returns the zoom details of the class at time.
    """
    zoom_class = zoom_meeting_id = zoom_password = zoom_day = zoom_time = zoom_link = ""

    zoom_details = [zoom_class, zoom_meeting_id, zoom_password, zoom_day, zoom_time, zoom_link]

    time_index_col = 4
    day_index_col = 3
    for row in range(sheet.nrows):
        zoom_time = sheet.cell_value(row,time_index_col)
        zoom_day = sheet.cell_value(row,day_index_col)
        if zoom_time == current_time and zoom_day == current_day:
            i = 0
            while i < len(zoom_details):
                zoom_details[i] = sheet.cell_value(row,i)
                i+=1
            return zoom_details

def set_clipboard():
    """
    Sets clipboard to zoom meeting id.
    """
    clipboard.copy(zoom_details[2])

def send_notification():
    """
    Sends zoom meeting password as notification on Mac.
    """

    title = "Sucessfully Joined Zoom Meeting!"
    message = f'The following zoom meeting password is copied to your clipboard: {zoom_details[2]}'
    command = f'''
    osascript -e 'display notification "{message}" with title "{title}"'
    '''
    os.system(command)

def on_fail_notification():
    """
    Sends notification on scheduling failure.
    """

    title = "Scheduling failed!"
    message = f'Please try again. Scheduling failed.'
    command = f'''
    osascript -e 'display notification "{message}" with title "{title}"'
    '''
    os.system(command)

def on_class_exists():
    """
    Function is called when there is a class at CURRENT_TIME
    """
    webbrowser.open(zoom_details[5], new=2)
    set_clipboard()
    send_notification()
    print("LOG >> Class at current time")

def on_class_none():
    """
    Function is called when there isn't a class at CURRENT_TIME
    """
    print("LOG >> No class at current time\n")


CURRENT_WEEKDAY = get_weekday_time()[0]
CURRENT_TIME = get_weekday_time()[1]
print("[CURRENT_WEEKDAY, CURRENT_TIME] >> ", get_weekday_time())

zoom_details = get_zoom_details(CURRENT_TIME,CURRENT_WEEKDAY)
print("ZOOM_DETAILS >> ", zoom_details)

if zoom_details != None:
    on_class_exists()
else:
    on_class_none()

########################## CRONTAB #############################

CRON_COMMAND = "/Users/peter/Desktop/Programming/Python/autoZoom/venv/bin/python3 /Users/peter/Desktop/Programming/Python/autoZoom/main.py >> /Users/peter/Desktop/Programming/Python/autoZoom/Logs.txt 2>&1"

def convert_hour(hour):
    if hour == "24":
        return "00"
    return hour


def check_day(day_index):
    days = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
    for day in days:
        if day.lower() == day_index.lower():
            return day
    on_fail_notification()

def convert_time_to_tab(time, day_index):
    minute = time[time.find(":")+1:]
    hour = convert_hour(time[:time.find(":")])
    day = check_day(day_index)
    return "{} {} * * {}".format(minute,hour,day)

cron = CronTab(user='peter')
cron.remove_all(comment='autoZoom')

time_index_col = 4
day_index_col = 3
for row in range(1, sheet.nrows):
    zoom_time = sheet.cell_value(row,time_index_col)
    zoom_day = sheet.cell_value(row,day_index_col)
    CRON_TAB = convert_time_to_tab(zoom_time,zoom_day)
    
    job = cron.new(command=CRON_COMMAND, comment="autoZoom")
    job.setall(CRON_TAB)

cron.write()

# for line in cron.lines:
#     print(line)