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
    return [DAYS[date.today().weekday()], get_utc_time()]

def get_utc_time():
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

def check_times(zoom_time, zoom_day):
    current = datetime.today()
    plus_early = current + timedelta(minutes=settings[JOIN_EARLY])
    b = (zoom_time == plus_early.strftime("%-H:%M")) and (zoom_day == DAYS[plus_early.weekday()])
    return b

def get_zoom_details():
    """
    Returns the zoom details of the class at time.
    """
    zoom_class = zoom_meeting_id = zoom_password = zoom_day = zoom_time = zoom_link = ""

    zoom_details = [zoom_class, zoom_meeting_id, zoom_password, zoom_day, zoom_time, zoom_link]

    for row in range(2, sheet.nrows):
        zoom_time = sheet.cell_value(row,TIME)
        zoom_day = sheet.cell_value(row,DAY)
        if check_times(zoom_time,zoom_day):
            i = 0
            while i < len(zoom_details):
                cv = sheet.cell_value(row,i)
                if cv == "":
                    cv = "NONE"
                zoom_details[i] = cv
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

def on_no_link_notification():
    """
    Sends notification when there is no link.
    """

    title = "Scheduling failed!"
    message = f'There is no link for: {zoom_details[CLASS]}'
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

def on_no_link():
    on_no_link_notification()
    print("LOG >> No link for the current class\n")

print("SHORT_DATE >> ", get_weekday_time())
print("FULL_DATE >> ", get_date())

zoom_details = get_zoom_details()
print("ZOOM_DETAILS >> ", zoom_details)

if zoom_details != None:
    if zoom_details[LINK] != "NONE":
        on_class_exists()
    else:
        on_no_link_notification()
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

    d = datetime(date.today().year,date.today().month,date.today().day,int(hours),int(minutes))
    dd = DAYS_TO_INTS[days.upper()]
    days_ahead = dd - d.weekday()

    if days_ahead <= 0:
        days_ahead += 7
    early_time = d + timedelta(days= days_ahead) - timedelta(minutes = int(settings[JOIN_EARLY]))

    return early_time.strftime("%H:%M:%a").split(":")

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