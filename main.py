import os
import webbrowser
from crontab import CronTab
import xlrd
from datetime import date, datetime, timedelta, timezone
import clipboard
from constants import *

def delete_logs():
    f = open(LOGS,"r+")
    if len(f.readlines()) > 5:
        f.truncate(0)
    f.close()

delete_logs()

schedule = xlrd.open_workbook(FILE_NAME)
sheet = schedule.sheet_by_index(0)
print(" ")

def get_date():
    x = datetime.today()
    return x.strftime("['%A, %b %d, %Y', '%H:%M']")

def get_weekday_time() -> list:
    """
    Returns the weekday and time in a list in the format [weekday, time]
    Weekday is in the list DAYS

    >>> get_weekday_time()
    ["MON", "12:01pm"]
    >>> get_weekday_time()
    ["WED", "11:53am"]
    """
    return [DAYS[date.today().weekday()], check_time()]

def check_time():
    """
    Returns current UTC time
    """
    current_utc_time = datetime.now().strftime("%H:%M")
    return current_utc_time

def parse_setting(s) -> str:
    return s[s.find("(") + 1:s.find(")")]

def get_settings():
    zoom_settings = []
    i = 1
    while i < 5:
        setting_str = parse_setting(sheet.cell_value(0, i))

        setting = setting_str
        try:
            setting = int(setting_str)
        except:
            if setting_str == "True":
                setting = True
            elif setting_str == "False":
                setting = False

        zoom_settings.append(setting)
        i+=1
    return zoom_settings
    
settings = get_settings()

def get_zoom_details(current_time, current_day):
    """
    Returns the zoom details of the class at time.
    """
    zoom_class = zoom_meeting_id = zoom_password = zoom_day = zoom_time = zoom_link = ""

    zoom_details = [zoom_class, zoom_meeting_id, zoom_password, zoom_day, zoom_time, zoom_link]

    for row in range(1, sheet.nrows):
        zoom_time = sheet.cell_value(row,TIME)
        zoom_day = sheet.cell_value(row,DAY)
        if zoom_time == current_time and zoom_day == current_day:
            i = 0
            while i < len(zoom_details):
                zoom_details[i] = sheet.cell_value(row,i)
                i+=1
            return zoom_details
        return None

def set_clipboard():
    """
    Sets clipboard to zoom meeting id.
    """
    clipboard.copy(zoom_details[ID])

def send_notification():
    """
    Sends zoom meeting password as notification on Mac.
    """

    title = "Sucessfully Joined Zoom Meeting!"
    message = f'The following zoom meeting password is copied to your clipboard: {zoom_details[PASSWORD]}'
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
    webbrowser.open(zoom_details[LINK], new=2)
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
print("SHORT_DATE >> ", get_weekday_time())
print("FULL_DATE >> ", get_date())

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
    for day in DAYS:
        if day.lower() == day_index.lower():
            return day
    on_fail_notification()

def get_early_time(days,hours,minutes):
    d = datetime(2020,1,date.today().day,int(hours),int(minutes))
    td = timedelta(minutes = int(settings[JOIN_EARLY]))

    original_time = d.strftime("%H:%M:%a").split(":")
    # print(original_time)
    early_time = (d - td).strftime("%H:%M:%a").split(":")
    # print(early_time)
    return early_time

def convert_time_to_tab(time, day_index):
    day = check_day(day_index)
    hour = convert_hour(time[:time.find(":")])
    minute = time[time.find(":")+1:]

    hour,minute,day = get_early_time(day,hour,minute)
    return "{} {} * * {}".format(minute,hour,check_day(day))

def print_crontab(cron):
    for line in cron.lines:
        print(line)

def set_crontab():
    cron = CronTab(user='peter')
    cron.remove_all(comment='autoZoom')

    if settings[TOGGLE_CRON]:
        for row in range(2, sheet.nrows):
            zoom_time = sheet.cell_value(row,TIME)
            zoom_day = sheet.cell_value(row,DAY)
            CRON_TAB = convert_time_to_tab(zoom_time,zoom_day)
            
            job = cron.new(command=CRON_COMMAND, comment="autoZoom")
            job.setall(CRON_TAB)

    cron.write()
    
    if settings[DEBUG]:
        print(("#" * 35) + " CRONTAB " + ("#" * 35))
        print_crontab(cron)


if __name__ == '__main__':
    set_crontab()