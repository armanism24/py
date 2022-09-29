# ----------( NOTES )-------------------------------------------------------------------------------
# ---[ 1 ]--| This script is written in Python-3
# ---[ 2 ]--| Modules/Drivers Included:
# ---------------[ 2.1 ]--| selenium            (Install by >> pip3 install selenium)
# ---------------[ 2.2 ]--| gspread             (Install by >> pip3 install gspread)
# ---------------[ 2.3 ]--| chromedriver.exe    (Download according to Chrome Version)
# ---[ 3 ]--| It wil generate log in 'log.txt' file.
# --------------------------------------------------------------------------------------------------
# ----------( SIGNATURE )---------------------------------------------------------------------------
# -----| Developer      :   Syed Qasim Ali Shah
# -----| Email          :   syedqasimalishah@protonmail.com
# -----| LinkedIn       :   https://www.linkedin.com/in/qasim-ecommerce-shopify-expert/
# --------------------------------------------------------------------------------------------------


# ----------( import modules )----------------------------------------------------------------------
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from datetime import datetime
from time import sleep
import gspread


# ----------( variables )---------------------------------------------------------------------------
SERVICE_ACCOUNT_FILE_PATH = r'Current\service_account.json'         # Enter path of JSON file download from Google Developer Account
WEBDRIVER_SERVICE_PATH = Service(r'Current\chromedriver.exe')       # Enter path of CHROMEDRIVER.EXE
LOG_FILE_PATH = r'Current\log.txt'                                  # Enter path of LOG.TXT (Text file)

G_SHEET_FILE_NAME = '211 Timesheet (Editable)'
WEBSITE_URL = 'https://secure5.saashr.com/ta/KS1367.clock'
# G_SHEET_LINK = 'https://docs.google.com/spreadsheets/d/1MyoGCsQNHiaiaarZ0AKjFmttouRFV3bnjEORZRQHOvA/edit?usp=sharing'

USERNAME = 'chernandez'                 # <----- Enter your USERNAME here
PASSWORD = 'S_q@s1m_password'           # <----- Enter your PASSWORD here
DEFAULT_WAIT = 600                      # <----- Enter the DEFAULT WAIT TIME here (i.e. 60s x 10 = 600s = 10 mins)

USERNAME_NAME = 'Username'
PASSWORD_NAME = 'Password'
CLOCK_IN_NAME = 'ChangeCC'
CLOCK_OUT_NAME = 'PunchOut'
POPUP_FRAME_ID = 'PopupBodyFrame'
PROJECTS_LIST_CLASS = 'treeNode'

next_task_no = 1
is_clocked_in = False


# ----------( access g-sheet )----------------------------------------------------------------------
print('-----| Welcome to the Bot |-----')
try:
    service_account = gspread.service_account(filename=SERVICE_ACCOUNT_FILE_PATH)
    worksheet = service_account.open(G_SHEET_FILE_NAME).sheet1
    print('Google Sheet is connected successfully!')
except:
    print('Google Sheet is not connected!')
    quit()


# ----------( functions )---------------------------------------------------------------------------
def get_day_of_year():
    '''
    Returns the day of year
    Return Type: <class 'int'>
    '''
    timestamp = datetime.now()
    return int(timestamp.strftime('%j'))

def get_datetime(year: int, date: str, time: str):
    '''
    Input : Year (e.g. 2022), Date (e.g. Jan 14), Time (e.g. 1:14 PM)
    Return : Timestamp <class 'datetime' | None>
    '''
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    if time != None:
        month_day = date.split()
        month = months.index(month_day[0]) + 1
        day = int(month_day[1])
        time_obj = datetime.strptime(time, "%I:%M %p")
        hour = int(datetime.strftime(time_obj, "%H"))
        minute = int(datetime.strftime(time_obj, "%M"))
        return datetime(year=year, month=month, day=day, hour=hour, minute=minute, second=0)
    else:
        return None

def get_schedule(day: int = get_day_of_year(), sheet=worksheet):
    '''
    Input : Day of Year <class 'int'>
    Return : Schedule <class 'dict'>
    Dictionary Keys : 'time_in_1-4', 'time_out_1-4', 'project_code_1-4'
    '''
    row_index = day + 1
    date = sheet.cell(row_index, 2).value
    year = datetime.now().year
    return {
        'time_in_1': get_datetime(year, date, sheet.cell(row_index, 3).value),
        'time_in_2': get_datetime(year, date, sheet.cell(row_index, 5).value),
        'time_in_3': get_datetime(year, date, sheet.cell(row_index, 7).value),
        'time_in_4': get_datetime(year, date, sheet.cell(row_index, 9).value),
        'time_out_1': get_datetime(year, date, sheet.cell(row_index, 4).value),
        'time_out_2': get_datetime(year, date, sheet.cell(row_index, 6).value),
        'time_out_3': get_datetime(year, date, sheet.cell(row_index, 8).value),
        'time_out_4': get_datetime(year, date, sheet.cell(row_index, 10).value),
        'project_code_1': sheet.cell(row_index, 11).value,
        'project_code_2': sheet.cell(row_index, 12).value,
        'project_code_3': sheet.cell(row_index, 13).value,
        'project_code_4': sheet.cell(row_index, 14).value
    }

def get_next_wait_time(timestamp: datetime):
    '''
    Input : Timestamp <class 'datetime'>
    Return : Wait Time in Seconds <class 'int'>
    '''
    current_timestamp = datetime.now()
    time_span = timestamp - current_timestamp
    return int(time_span.total_seconds())

def print_log(text: str):
    '''
    Input : Text <class 'str'>
    Function : Print on screen and Store log in 'log.txt' file
    '''
    current_time = str(datetime.now())
    formatted_time = current_time[:current_time.rfind('.')] + ' : '
    with open(LOG_FILE_PATH, 'a', encoding='UTF-8') as file:
        file.write(formatted_time + text + '\n')
    print(formatted_time + text)


# ----------( open chrome browser )-----------------------------------------------------------------
print_log('Website loading procedure has been initiated!')
driver = webdriver.Chrome(service=WEBDRIVER_SERVICE_PATH)
driver.maximize_window()
driver.get(WEBSITE_URL)
driver.implicitly_wait(5)
print_log('Website loaded successfully!')
sleep(10)


# ----------( automation script )-------------------------------------------------------------------
while True:
    while True:
        schedule = get_schedule()
        print_log('Scedule = ' + str(schedule))
        wait_time = 0

        if schedule['time_in_{}'.format(next_task_no)] == None:
            print_log('Task # {} not found! Waiting for {}s'.format(next_task_no, DEFAULT_WAIT))
            sleep(DEFAULT_WAIT)
            # next_task_no += 1
            # if next_task_no > 4:
            #     next_task_no = 1
            #     sleep(DEFAULT_WAIT)
            continue

        if is_clocked_in == False:
            wait_time = get_next_wait_time(schedule['time_in_{}'.format(next_task_no)])
            print_log('Waiting time calculated of Clock-In for Task # {}'.format(next_task_no))
        else:   # elif is_clocked_in == True:
            wait_time = get_next_wait_time(schedule['time_out_{}'.format(next_task_no)])
            print_log('Waiting time calculated of Clock-Out for Task # {}'.format(next_task_no))
        
        if wait_time < 0:
            print_log('Waiting time is negative (i.e. {}s)'.format(wait_time))
            next_task_no += 1
            if next_task_no > 4:
                next_task_no = 1
                sleep(DEFAULT_WAIT)
            continue
        elif wait_time > DEFAULT_WAIT:
            print_log('Waiting time is greater than {}s (i.e. {}s). Waiting...'.format(DEFAULT_WAIT, wait_time))
            sleep(DEFAULT_WAIT)
            continue
        else:
            print_log('Waiting time is less than {}s (i.e. {}s). Waiting...'.format(DEFAULT_WAIT, wait_time))
            sleep(wait_time)
            break

    try:
        username_input = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, USERNAME_NAME)))
        username_input.clear()
        username_input.send_keys(USERNAME)
        password_input = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, PASSWORD_NAME)))
        password_input.clear()
        password_input.send_keys(PASSWORD)
        print_log('Username and Password has been entered!')
    except Exception as E:
        print(E)

    current_project = schedule['project_code_{}'.format(next_task_no)]
    print_log('Current Project : {}'.format(current_project))

    if is_clocked_in == False:
        try:
            clock_in_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.NAME, CLOCK_IN_NAME)))
            clock_in_btn.click()
            popup_frame = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, POPUP_FRAME_ID)))
            driver.switch_to.frame(popup_frame)
            projects_list = WebDriverWait(driver, 5).until(EC.presence_of_all_elements_located((By.CLASS_NAME, PROJECTS_LIST_CLASS)))
            for project in projects_list:
                if project.text.strip() == current_project.strip():
                    # project.click()
                    print_log('Clock-In SUCCESSFULL!')
            driver.switch_to.default_content()
        except Exception as E:
            print(E)
        is_clocked_in = True
    elif is_clocked_in == True:
        try:
            clock_out_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.NAME, CLOCK_OUT_NAME)))
            # clock_out_btn.click()
            print_log('Clock-Out SUCCESSFULL!')
        except Exception as E:
            print(E)
        is_clocked_in = False
        next_task_no += 1
        if next_task_no > 4:
            next_task_no = 1


driver.quit()
