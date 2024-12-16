import os
import gspread
import sys
import traceback


def authenticate(auth):
    try:
        googleService = gspread.service_account(filename=auth)
        return googleService
    except FileNotFoundError:
        print("Error: Json file not found")
    except Exception as e:
        print("Error: Authentication failed")
        print(e)


def getPath():
    try:
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # If so, use the bundle's directory
            base_path = sys._MEIPASS
        else:
            # Otherwise, use the directory of the script (or executable)
            base_path = os.path.dirname(__file__)
    except:
        base_path = os.getcwd()
        print('Error getting base path:', traceback.format_exc())
    return base_path


def getSheet(planilha):
    basePath = getPath()
    # jsonFilePath = 'credentials.json'
    jsonFilePath = os.path.join(basePath, 'credentials.json')
    googleService = authenticate(jsonFilePath)
    sheet: gspread.Spreadsheet = googleService.open_by_url(planilha)
    return sheet

