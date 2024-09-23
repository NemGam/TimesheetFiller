#region imports
import os
from pathlib import Path
from win32com import client
import pywintypes

#Change to https://stackoverflow.com/questions/1180115/add-text-to-existing-pdf-using-python
import json #Maybe use TOML instead
import datetime

#To work with calendar
from icalendar import Calendar
import openpyxl.workbook
import recurring_ical_events
import requests

#To work with Excel
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
#endregion

FULL_NAME : str = ""
EVENT_NAME : str = ""
ICAL_LINK : str = ""

CURRENT_DIRECTORY = Path(os.getcwd())
EXCEL_FILE = CURRENT_DIRECTORY.joinpath("Timesheet.xlsx")

    
def load_data():
    '''Loads data from config.json or gets it from user'''
    global FULL_NAME
    global EVENT_NAME
    global ICAL_LINK

    save_path = CURRENT_DIRECTORY.joinpath("config.json")

    if Path(save_path).exists(): #File found
        config = open(save_path, 'r')
        data = json.loads(config.read())
        FULL_NAME = data["name"]
        EVENT_NAME = data["event"]
        ICAL_LINK = data["link"]
        config.close()
    else: #File not found
        config = open(save_path, 'w')
        FULL_NAME = input("Enter your FULL NAME: ")
        EVENT_NAME = input("Enter the name of the event you want to read (as it appears in calendar): ")
        ICAL_LINK = input("Enter the ICal link from your calendar: ")
        data = {
            "name": FULL_NAME,
            "event": EVENT_NAME,
            "link": ICAL_LINK
        }
        json.dump(data, config)
        config.close()


def get_current_pay_period(wb : openpyxl.Workbook) -> Worksheet:
    '''Calculates current pay period using current date and sheets names'''    
    input_date = datetime.date.today()
    final_sheet : Worksheet = None
    for sheet in wb:
        sheet.views.sheetView[0].tabSelected = False #Fixes bug that causes multiple tabs to be active at the same time
        time = 0
        try:
            time = datetime.datetime.strptime(sheet.title, "%m_%d_%Y")
        except ValueError:
            time = datetime.datetime.strptime(sheet.title, "%m_%d_%y")
        
        if not final_sheet and time.date() >= input_date:
            final_sheet = sheet
    
    if final_sheet is not None:
        return final_sheet
    
    raise Exception("The pay period was not found!")      


def calculate_PP_dates(sample_date : str) -> list[datetime.datetime]:
    '''
    Calculates start and end dates of the pay period
    '''
    date = 0
    try:
        date = datetime.datetime.strptime(sample_date, '%m_%d_%Y')
    except ValueError:
        date = datetime.datetime.strptime(sample_date, "%m_%d_%y")
    s = date + datetime.timedelta(days= 1)
    s1 = date + datetime.timedelta(days= -13)

    #Filling dates dictionary
    calculate_two_weeks(date)
    return [s1, s]


def calculate_two_weeks(date : datetime.datetime):
    '''Calculates dates of current pay period'''
    for _ in range(14):
        dates[date.isoformat()[:10]] = []
        date += datetime.timedelta(days= -1)


def process_events(eventsToFill):
    '''Separates date, start and end times of given events'''
    for event in eventsToFill:
        date = event[0].date().strftime("%Y-%m-%d")
        start = event[0].time()
        end = event[1].time()

        event[0] = start
        event[1] = end
        event.insert(0, date)
        dates[date].append([start, end])


def fill_excel():
    '''Fills excel with known dates'''
    current_row : int = 35
    sheet = wb.active
    if sheet == None:
        raise Exception("Active worksheet not found!")
    
    if wb.active != None:

        for key in dates:
            ls = list(dates.get(key))
            ls.reverse()
            count = 0
            current_column : int = ord("D")
            for times in ls:
                if count < 3:
                    count += 1
                    sheet[f"{chr(current_column)}{current_row}"] = times[0]
                    sheet[f"{chr(current_column + 1)}{current_row}"] = times[1]
                    current_column += 2
                else:
                    #TODO:Implement events connection (1 PM-3 PM + 3 PM-5 PM = 1 PM-5PM)
                    print("WARNING, YOU HAVE MORE THAN 3 EVENTS IN THE SAME DAY. PLEASE FILL THIS DAY MANUALLY!")
                    break
            
            
            if current_row == 29:  #Jump through the gap in the timesheet
                current_row -= 10
            else:
                current_row -= 1
        
    else:
        raise Exception("Active list is not found!")   


def parse_calendar(start_date, end_date):
    '''Parses calendar and returns events between given dates'''

    events_to_fill = []
    cal = Calendar.from_ical(requests.get(ICAL_LINK).text)
    l = recurring_ical_events.of(cal).between(start_date, end_date)
    for event in l:
        if event.get("summary") != EVENT_NAME:
            continue
        events_to_fill.append([event.get("DTSTART").dt, event.get("DTEND").dt])
    
    return events_to_fill


def convert_to_PDF(sheet_to_convert_ID : int):
    '''
    Converts filled Excel file to PDF
    '''
    
    try:
        #Launch Excel
        excel = client.Dispatch("Excel.Application")

        # Read Excel File
        sheets = excel.Workbooks.Open(f"{CURRENT_DIRECTORY.joinpath(f'{FULL_NAME} {activeSheet.title}.xlsx')}")
        work_sheets = sheets.Worksheets[sheet_to_convert_ID]
        
        # Convert into PDF File
        work_sheets.ExportAsFixedFormat(0, f"{CURRENT_DIRECTORY.joinpath(f'{FULL_NAME} {activeSheet.title}.pdf')}")
        sheets.Close(False)
        print(f"Converted successfully! ({CURRENT_DIRECTORY.joinpath(f'{FULL_NAME} {activeSheet.title}.pdf')})")
    
    except pywintypes.com_error as error:
        print(error.args[2][2]) #Message
        print("No PDF conversion happened.")
    
    except Exception as e:
        print("Error occured while converting! " + e)
        print("No PDF conversion happened.")

    finally:
        #Don't forget to terminate Excel
        excel.Quit()
    



if __name__ == '__main__':

    print("Loading data...")
    load_data()

    dates = {}
    wb : openpyxl.Workbook

    #Open workbook
    try:
        wb = openpyxl.load_workbook(filename=EXCEL_FILE)
    except FileNotFoundError as error:
        print(error)
        print("Please place 'Timesheet.xlsx' in the same directory as this file.")
        input("Press Enter to exit...")
        os.abort()

    print("Working...")
    
    wb.calculation.calcMode = 'auto'
    active_sheet_index : int = 0
    activeSheet = get_current_pay_period(wb)
    if activeSheet == None:
        print("Active workbook not found!")
        input("Press Enter to exit...")
        os.abort()

    wb.active = activeSheet
    active_sheet_index = wb.index(activeSheet)

    startDate, endDate = calculate_PP_dates(activeSheet.title)
    eventsToFill = parse_calendar(startDate, endDate)
    process_events(eventsToFill)

    fill_excel()
    
    wb.save(CURRENT_DIRECTORY.joinpath(f"{FULL_NAME} {activeSheet.title}.xlsx"))
    wb.close()

    print("Converting to PDF...")
    convert_to_PDF(active_sheet_index)
    
    print(f"Excel file successfully created. ({CURRENT_DIRECTORY.joinpath(f'{FULL_NAME} {activeSheet.title}.xlsx')})")
    input("Press Enter to exit...")
