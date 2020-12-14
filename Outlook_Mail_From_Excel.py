# -*- coding: utf-8 -*-
"""
Created on Fri Oct 30 10:07:23 2020

@author: Guilherme Bresciani de Azevedo
"""
#TO DO
#Try to format value got from excel based on 'CELL.NumberFormat' info. 

#Print SW version
print("| RTA ----------------------- DEA-M ----------------------- ALGO |")
print("| Tool: Outlook_Mail_From_Excel |        Version:  V01.10        |")
print("|----------------------------------------------------------------|",
      end='\n\n')

import tkinter as tk
import os
import time
import sys
from tkinter import filedialog
import re
import logging
from collections.abc import Iterable
from PIL import ImageGrab
import win32com.client as client

#Class definition
class SafeDict(dict):
    def __missing__(self, key):
        return '{' + key + '}'

#Function defition
def safe_exit(to_Close= None, to_Quit=None, handler=None, app_root=None, opt_exit=True):
    logger.info("Closing...")
    if isinstance(to_Close, Iterable):
        for item in to_Close:
            try:
                item.Close()
            except:
                pass
    elif to_Close is not None:
        try:
            to_Close.Close()
        except:
            pass
    if isinstance(to_Quit, Iterable):
        for item in to_Quit:
            try:
                item.Quit()
            except:
                pass
    elif to_Quit is not None:
        try:
            to_Quit.Quit()
        except:
            pass
    if isinstance(handler, Iterable):
        for hnd in handler:
            try:
                hnd.flush()
                hnd.close()
                logger.removeHandler(hnd)
            except:
                pass
    elif handler is not None:
        try:
            handler.flush()
            handler.close()
            logger.removeHandler(handler)
        except:
            pass
    if app_root is not None:
        try:
            app_root.destroy()
        except:
            pass
    if opt_exit:
        sys.exit()
        
def ask_for_files(diag_parent, diag_title, file_types=[('All files','*.*')], at_least_one=True, more_than_one=False, second_title=''):
    if second_title == '':
        second_title = diag_title
    PATH_LIST = []
    PATH = "first time asking"
    while len(PATH) > 0:
        PATH = filedialog.askopenfilename(parent=diag_parent, title=diag_title, filetypes=file_types)
        if len(PATH) > 0: #user selected a file
            logger.info("User selected file path '{}'".format(PATH))
            if more_than_one: #should ask more
                PATH_LIST.append(PATH)
                diag_title = second_title
                next
            else: #1 file selected and only 1 required
                return PATH
        elif len(PATH_LIST) > 0: #enough files selected and more_than_one=True
            return PATH_LIST
        elif at_least_one: #user not selected at first ask and at least 1 file required
            logger.error("A file was not selected when asked '{}' !!!".format(diag_title))
            return None
        else:
            logger.info("A file was not selected when asked '{}'.".format(diag_title))
            return PATH_LIST
    
def get_list_from_txt_file(file_path, remove_header=False, list_headers=[], dict_format={}, raise_not_found=True):
    USER_LIST = []
    try: #Try to open the file
        with open(file_path) as FILE:
            USER_LIST = [line.strip() for line in FILE if line.strip()] #list by lines and format with dict
    except FileNotFoundError:
        if raise_not_found:
            logger.exception("File '{}' not found !!!".format(file_path))
            raise
        else:
            logger.info("Optional file '{}' not found.".format(file_path))
            return USER_LIST
    except Exception:
        logger.exception("Error reading file '{}' !!!".format(file_path))
        raise
    #Remove header from user list
    if remove_header and len(USER_LIST) > 0:
        if len(list_headers) == 0:
            USER_LIST.pop(0)
        else:
            for TITLE in list_headers:
                if USER_LIST[0].casefold() == TITLE.casefold() or USER_LIST[0].casefold() == TITLE.casefold() + 's':
                    USER_LIST.pop(0)
                    logger.info("Header '{}' removed from file '{}'".format(USER_LIST[0], file_path))
    #Check exitence of information
    if len(USER_LIST) == 0:
        logger.warning("File empty '{}' !".format(file_path))
    else:
        logger.info("File read in '{}'.".format(file_path))
        if len(dict_format) > 0: #Format using dict
            for INDEX, ITEM in enumerate(USER_LIST):
                try:
                    USER_LIST[INDEX] = ITEM.format_map(SafeDict(dict_format))
                except:
                    logger.warning("Error formating setup '{}' in file '{}' using manually setted values!!!".format(ITEM, file_path))
                    pass
                    
    return USER_LIST
    #File empty do not raise stop inplace, choose outside.
    
def get_index_by_str(idx_text, opt_list=[], less_one=True):
    if idx_text.isnumeric():
        if len(opt_list) == 0 or (len(opt_list) > 0 and len(opt_list) >= int(idx_text)):
            if less_one:
                return int(idx_text) - 1
            else:
                return int(idx_text)
    else:
        for idx, option in enumerate(opt_list):
            if idx_text in option: #if option contains the idx_text suggested
                return idx
    logger.warning("Index not found for user defined index '{}' inside options '{}' !".format(idx_text, opt_list))
    return None #in case of not finding a valid index

def find_column(sheet, coded_location):
    #coded_location example: (1;==;any text;+;0)
    SET_COL = -1
    for COLUMN in range(1,sheet.UsedRange.Columns.Count + 2): #+2 test until the first empty cell
        CELL_VALUE = sheet.Cells(int(coded_location.group('col_in_row')), COLUMN).Value
        if coded_location.group('col_comp') == "==":
            if CELL_VALUE == coded_location.group('col_text'):
                if coded_location.group('col_offset') == "+":
                    SET_COL = COLUMN + int(coded_location.group('col_off_value'))
                    break
                elif coded_location.group('col_offset') == "-":
                    SET_COL = COLUMN - int(coded_location.group('col_off_value'))
                    break
        elif coded_location.group('col_comp') == "!=" or coded_location.group('col_comp') == "<>":
            if CELL_VALUE != coded_location.group('col_text'):
                if coded_location.group('col_offset') == "+":
                    SET_COL = COLUMN + int(coded_location.group('col_off_value'))
                    break
                elif coded_location.group('col_offset') == "-":
                    SET_COL = COLUMN - int(coded_location.group('col_off_value'))
                    break
        elif coded_location.group('col_comp') == ">":
            if CELL_VALUE > int(coded_location.group('col_text')):
                if coded_location.group('col_offset') == "+":
                    SET_COL = COLUMN + int(coded_location.group('col_off_value'))
                    break
                elif coded_location.group('col_offset') == "-":
                    SET_COL = COLUMN - int(coded_location.group('col_off_value'))
                    break
        elif coded_location.group('col_comp') == "<":
            if CELL_VALUE < int(coded_location.group('col_text')):
                if coded_location.group('col_offset') == "+":
                    SET_COL = COLUMN + int(coded_location.group('col_off_value'))
                    break
                elif coded_location.group('col_offset') == "-":
                    SET_COL = COLUMN - int(coded_location.group('col_off_value'))
                    break
        else:
            next
    if SET_COL < 0:
        logger.warning("Column not found for logic '{}' !".format(coded_location.group('col')))
    else:
        logger.info("Column '{}' found for logic '{}'.".format(SET_COL, coded_location.group('col')))
        
    return SET_COL

def find_row(sheet, coded_location):
    #coded_location example: (A;==;any text;+;0)
    SET_ROW = -1
    for ROW in range(1,sheet.UsedRange.Rows.Count + 2): #+2 test until the first empty cell
        CELL_VALUE = sheet.Range(coded_location.group('lin_in_col') + str(ROW)).Value
        if coded_location.group('lin_comp') == "==":
            if CELL_VALUE == coded_location.group('lin_text'):
                if coded_location.group('lin_offset') == "+":
                    SET_ROW = ROW + int(coded_location.group('lin_off_value'))
                    break
                elif coded_location.group('lin_offset') == "-":
                    SET_ROW = ROW - int(coded_location.group('lin_off_value'))
                    break
        elif coded_location.group('lin_comp') == "!=" or coded_location.group('lin_comp') == "<>":
            if CELL_VALUE != coded_location.group('lin_text'):
                if coded_location.group('lin_offset') == "+":
                    SET_ROW = ROW + int(coded_location.group('lin_off_value'))
                    break
                elif coded_location.group('lin_offset') == "-":
                    SET_ROW = ROW - int(coded_location.group('lin_off_value'))
                    break
        elif coded_location.group('lin_comp') == ">":
            if CELL_VALUE > int(coded_location.group('lin_text')):
                if coded_location.group('lin_offset') == "+":
                    SET_ROW = ROW + int(coded_location.group('lin_off_value'))
                    break
                elif coded_location.group('lin_offset') == "-":
                    SET_ROW = ROW - int(coded_location.group('lin_off_value'))
                    break
        elif coded_location.group('lin_comp') == "<":
            if CELL_VALUE < int(coded_location.group('lin_text')):
                if coded_location.group('lin_offset') == "+":
                    SET_ROW = ROW + int(coded_location.group('lin_off_value'))
                    break
                elif coded_location.group('lin_offset') == "-":
                    SET_ROW = ROW - int(coded_location.group('lin_off_value'))
                    break
        else:
            next
    if SET_ROW < 0:
        logger.warning("Row not found for logic '{}' !".format(coded_location.group('lin')))
    else:
        logger.info("Row '{}' found for logic '{}'.".format(SET_ROW, coded_location.group('lin')))
    
    return SET_ROW

def get_cell(sheet, location):
    if location.group('col_in_row') is None and location.group('lin_in_col') is None: #location == A1
        return sheet.Range(location.group('cell')) #sheet.Range because Location is str
    elif location.group('col_in_row') is not None and location.group('lin_in_col') is None: #location == (logic)1
        COLUMN = find_column(sheet, location)
        if COLUMN < 0: #if not found
            logger.warning("Column not found for setup '{}' !".format(location.group(0)))
            return None
        else:
            return sheet.Cells(int(location.group('lin')), COLUMN) #sheet.Cells because Column is int
    elif location.group('col_in_row') is None and location.group('lin_in_col') is not None: #location == A(logic)
        ROW = find_row(sheet, location)
        if ROW < 0: #if not found
            logger.warning("Row not found for setup '{}' !".format(location.group(0)))
            return None
        else:
            return sheet.Range(location.group('col') + str(ROW))
    elif location.group('col_in_row') is not None and location.group('lin_in_col') is not None: #location == (logic)(logic)
        COLUMN = find_column(sheet, location)
        if COLUMN < 0: #if not found
            logger.warning("Column not found for setup '{}' !".format(location.group(0)))
            return None
        ROW = find_row(sheet, location)
        if ROW < 0: #if not found
            logger.warning("Row not found for setup '{}' !".format(location.group(0)))
            return None
        return sheet.Cells(ROW, COLUMN)
    #Cell not found do not raise error, only warn user and return None
    
def delete_file(file_path):
    if os.path.exists(file_path): #file exists
        os.remove(file_path) #delete file
        return True
    else:
        return False
############################################################

#Global declaration
TIME_TO_SLEEP = 5

#Window app declaration
APP_ROOT = tk.Tk()
APP_ROOT.attributes("-topmost", True) # to open dialogs in front of other windows
#APP_ROOT.lift()
APP_ROOT.withdraw() #hide application main window
try:
    APP_ROOT.iconbitmap(os.getcwd() + '\\icon.ico')
except:
    pass

#Get email Template .htm* file
HTML_FILE_PATH = filedialog.askopenfilename(parent=APP_ROOT, title="Select email template", 
                                           filetypes=[('HTML files',
                                                       ['*.htm', '*.html'])])
try: #try to read html file on path informed by user
    with open(HTML_FILE_PATH) as USER_FILE:
        HTML_BODY = USER_FILE.read() #read as one string
except:
    print("Error reading HTML file '{}' !!!".format(HTML_FILE_PATH))
    time.sleep(TIME_TO_SLEEP)
    sys.exit()    
if len(HTML_BODY) == 0:
    print("File empty '{}' !!!".format(HTML_FILE_PATH))
    time.sleep(TIME_TO_SLEEP)
    sys.exit()

#Define files folder
LIST_POSSIBLE_FOLDERS = ['_arquivos', '_files', '_file', '_fichiers']
FILES_DIR = ''
for POSSIBLE_FOLDER in LIST_POSSIBLE_FOLDERS:
    if os.path.exists(HTML_FILE_PATH.split('.htm')[0] + POSSIBLE_FOLDER): #folder exists in same path o HTML file
        FILES_DIR = HTML_FILE_PATH.split('.htm')[0] + POSSIBLE_FOLDER
        break
if FILES_DIR == '': #folder do not exists with expected names in same path of HTML file
    print("No file folder corresponding to file '{}' was found in path '{}' !!!".format('/'.join(HTML_FILE_PATH.split('/')[0:-1]), HTML_FILE_PATH.split('/')[-1]))
    time.sleep(TIME_TO_SLEEP)
    sys.exit()
        
#Log declaration
delete_file(FILES_DIR + '/' + 'log.log') #logging only last time used software

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

file_formatter = logging.Formatter('%(levelname)s : %(message)s')
file_handler = logging.FileHandler(FILES_DIR + '/' + 'log.log')
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(file_formatter)

stream_formatter = logging.Formatter('%(levelname)s : %(message)s')
stream_handler = logging.StreamHandler()
stream_handler.setLevel(logging.WARNING)
stream_handler.setFormatter(stream_formatter)

logger.addHandler(file_handler)
logger.addHandler(stream_handler)
logger.propagate = False

#Check Manual values setting file .txt
TXT_FILE_PATH = FILES_DIR + "/setting_value_manually.txt"
USER_VAL_MAN_LIST = get_list_from_txt_file(TXT_FILE_PATH, remove_header=True, raise_not_found=False)

#Start dict of value to replace
VAL_DICT = {}
for ID in USER_VAL_MAN_LIST: #IDs defined in setting_value_manually.txt
    USER_MAN_VALUE = str(input("Write a value for the ID tag named '{}':\n".format(ID))) #Ask value for key
    VAL_DICT[ID] = USER_MAN_VALUE

#Check Image setting file .txt
TXT_FILE_PATH = FILES_DIR + "/setting_image.txt"
USER_IMG_LIST = get_list_from_txt_file(TXT_FILE_PATH, remove_header=True, dict_format=VAL_DICT)

#Check Value setting file .txt
TXT_FILE_PATH = FILES_DIR + "/setting_value.txt"
USER_VAL_LIST = get_list_from_txt_file(TXT_FILE_PATH, remove_header=True, dict_format=VAL_DICT)
    
#Check Send To setting file.txt
TXT_FILE_PATH = FILES_DIR + "/setting_send_to.txt"
USER_SEND_TO_LIST = get_list_from_txt_file(TXT_FILE_PATH, remove_header=True, dict_format=VAL_DICT)

#Check Subject setting file.txt
TXT_FILE_PATH = FILES_DIR + "/setting_subject.txt"
USER_SUBJECT_LIST = get_list_from_txt_file(TXT_FILE_PATH, remove_header=True, dict_format=VAL_DICT)
USER_SUBJECT = USER_SUBJECT_LIST[0]
    
#Ask for Excel to get information from
DIALOG_TITLE = "Select one Excel file to get information"
DIALOG_2ND_TITLE = "Select another Excel file to get information or Cancel"
DIALOG_FILE_TYPES = [('Excel files', ['*.xlsx', '*.xlsx', '*.xlsm'])]
XLS_FILE_LIST = ask_for_files(APP_ROOT, DIALOG_TITLE, DIALOG_FILE_TYPES, 
                              more_than_one=True, second_title=DIALOG_2ND_TITLE)

#Check if at least one Excel file was selected      
if XLS_FILE_LIST == None: #user closed the first ask dialog
    safe_exit(handler=[file_handler, stream_handler], app_root=APP_ROOT) #exit and logging intentional exit
#List names of Worsheets for indexing data collection by name purposes
XLS_FILE_NAMES = [FILE_PATH.split('/')[-1] for FILE_PATH in XLS_FILE_LIST]

######################## ↓ EXCEL ↓ ########################
#New instance of Excel
EXCEL = client.DispatchEx('Excel.Application')
logger.info("Opened Excel application.")
#EXCEL.ScreenUpdating = False
#EXCEL.DisplayAlerts = False
#EXCEL.Visible = False

#Open workbooks
WORKBOOK = []
for PATH in XLS_FILE_LIST:
    try:
        WORKBOOK.append(EXCEL.Workbooks.Open(PATH))
        logger.info("Opened Workbook '{}'.".format(PATH))
    except:
        logger.error("Error opening the Workbook '{}' !!!".format(PATH))
        safe_exit(to_Close=WORKBOOK, to_Quit=EXCEL, handler=[file_handler, stream_handler], app_root=APP_ROOT)

#List worksheets in same index order of WORKBOOK list
SHEET_LISTS = []
for BOOK in WORKBOOK:
    SHEET_LIST = []
    for SHEET in BOOK.Worksheets:
        SHEET_LIST.append(SHEET.Name)
    SHEET_LISTS.append(SHEET_LIST)
logger.info("Listed Worsheets by Workbook '{}'.".format(SHEET_LISTS))

#Set pattern for user setup of 'cell location'
PATTERN = re.compile(r"""
                     (?P<cell>
                     (?P<col> [A-Z]+ |
                     \(
                     (?P<col_in_row> \d+);
                     (?P<col_comp> == | != | <> | >= | <=);
                     (?P<col_text> .*);
                     (?P<col_offset> \+ | -);
                     (?P<col_off_value> \d+)
                     \))
                     (?P<lin> \d+ | 
                     \(
                     (?P<lin_in_col> [A-Z]+);
                     (?P<lin_comp> == | != | <> | >= | <=);
                     (?P<lin_text> .*);
                     (?P<lin_offset> \+ | -);
                     (?P<lin_off_value> \d+)
                     \)))""", flags=re.VERBOSE | re.DOTALL)
                     
######################## ↓ GET VALUES ↓ ########################
for SETUP in USER_VAL_LIST: #SETUP = "ID	Sheet	Cell"
    SETUP = SETUP.format_map(SafeDict(VAL_DICT)) #format setup with values previously got
    KEY = str(SETUP.split(sep='\t')[0]) #dict key to format_map
    BOOK_INDEX = get_index_by_str(SETUP.split(sep='\t')[1], XLS_FILE_NAMES)
    if BOOK_INDEX == None:
        logger.warning("Could not found Workbook for setup '{}' !".format(SETUP))
        next
    BOOK = WORKBOOK[BOOK_INDEX] #user Book=1 will access by index 0
    SHEET_INDEX = get_index_by_str(SETUP.split(sep='\t')[2], SHEET_LISTS[BOOK_INDEX])
    if SHEET_INDEX == None:
        logger.warning("Could not found Worksheet for setup '{}' !".format(SETUP))
        next
    SHEET = BOOK.Sheets[SHEET_INDEX] #user Sheet=1 will access by index 0
    LOCATION = re.match(PATTERN, str(SETUP.split(sep='\t')[3])) #recognize pattern in user defined location
    if LOCATION is not None: #location could be recognized as standard
        CELL = get_cell(SHEET, LOCATION) #decode LOCATION if logic present
        if CELL is not None: #Successfully got cell
            try: #try to get .Value attribute of cell
                VAL_DICT[KEY] = str(CELL.Value)
            except:
                logger.warning("Value not found for Cell in setup '{}' !".format(SETUP))
        else:
            logger.warning("Cell not found for Value setup '{}' !".format(SETUP))
    else: #user defined location not recognizable
        logger.warning("Standard not recognized for Value setup '{}' !".format(SETUP))
        #Values not found do not raise error, only warn user and remain as empty string
#Register values found to user
logger.info("Values found for replacement are '{}'.".format(VAL_DICT))
print("\nValues found for replacement are:", end='\n')
[print(KEY + ": " + VAL_DICT[KEY]) for KEY in VAL_DICT]
print("\n")
######################## ↑ GET VALUES ↑ ########################

######################## ↓ GET IMAGES ↓ ########################
for SETUP in USER_IMG_LIST: #SETUP = "ID	Sheet	Type	Cell/Num"
    SETUP = SETUP.format_map(SafeDict(VAL_DICT)) #format image setup with values got
    IMAGE_ID = SETUP.split(sep='\t')[0] #image file name
    BOOK_INDEX = get_index_by_str(SETUP.split(sep='\t')[1], XLS_FILE_NAMES)
    if BOOK_INDEX == None:
        logger.warning("Could not found Workbook for setup '{}' !".format(SETUP))
        delete_file(FILES_DIR + '/' + IMAGE_ID) #delete wrong old image 
        next
    BOOK = WORKBOOK[BOOK_INDEX] #user Book=1 will access by index 0
    SHEET_INDEX = get_index_by_str(SETUP.split(sep='\t')[2], SHEET_LISTS[BOOK_INDEX])
    if SHEET_INDEX == None:
        logger.warning("Could not found Worksheet for setup '{}' !".format(SETUP))
        delete_file(FILES_DIR + '/' + IMAGE_ID) #delete wrong old image 
        next
    SHEET = BOOK.Sheets[SHEET_INDEX] #user Sheet=1 will access by index 0
    TYPE = SETUP.split(sep='\t')[3] #Type = 'table' or 'chart'
    LOCATION = SETUP.split(sep='\t')[4]
    if TYPE == 'table':
        BGN_LOCATION = str(LOCATION.split(sep=':')[0]) #get A1 from range A1:B2
        END_LOCATION = str(LOCATION.split(sep=':')[1]) #get B2 from range A1:B2
        BGN_LOCATION = re.match(PATTERN, BGN_LOCATION) #recognize user defined location
        END_LOCATION = re.match(PATTERN, END_LOCATION) #recognize user defined location
        if BGN_LOCATION is not None and END_LOCATION is not None: #location could be recognized as standard
            BGN_CELL = get_cell(SHEET, BGN_LOCATION) #decode LOCATION if logic present
            END_CELL = get_cell(SHEET, END_LOCATION) #decode LOCATION if logic present
            if BGN_CELL is not None and END_CELL is not None: #location could be recognized as standard
                try: #try to select and copy range
                    RANGE = SHEET.Range(BGN_CELL, END_CELL)
                    RANGE.CopyPicture(Appearance=1, Format=2)
                    try: #try to save image copied
                        ImageGrab.grabclipboard().save(FILES_DIR + '/' + IMAGE_ID)
                    except:
                        logger.warning("Could not save image for setup '{}' !".format(SETUP))
                        delete_file(FILES_DIR + '/' + IMAGE_ID) #delete wrong old image 
                        next
                except:
                    logger.warning("Range not found for setup '{}' !".format(SETUP))
                    delete_file(FILES_DIR + '/' + IMAGE_ID) #delete wrong old image
                    next
            else:
                logger.warning("Cell not found for Value setup '{}' !".format(SETUP))
                delete_file(FILES_DIR + '/' + IMAGE_ID) #delete wrong old image 
                next
        else: #user defined location not recognizable
            logger.warning("Standard not recognized for Value setup '{}' !".format(SETUP))
            delete_file(FILES_DIR + '/' + IMAGE_ID) #delete wrong old image  
            next
    elif TYPE == 'chart':
        try:
            CHART = SHEET.ChartObjects(int(LOCATION)) #select chart by number
            CHART.Activate() #Avoid exporting an image with nothing inside
            try:
                CHART.Chart.Export(FILES_DIR + '/' + IMAGE_ID)            
            except:
                logger.warning("Could not save image for setup '{}' !".format(SETUP))
                delete_file(FILES_DIR + '/' + IMAGE_ID) #delete wrong old image
                next
        except:
            logger.warning("Chart no reachable for setup '{}' !".format(SETUP))
            delete_file(FILES_DIR + '/' + IMAGE_ID) #delete wrong old image
            next
    else: #user defined type not recognizable
        logger.warning("Type not recognized for setup '{}' !".format(SETUP))
        delete_file(FILES_DIR + '/' + IMAGE_ID) #delete wrong old image
        next
#Image not reachable do not raise error, only warn user and assure no wrong image     
######################## ↑ GET IMAGES ↑ ########################

#Close Workbooks and Excel
safe_exit(to_Close=WORKBOOK, to_Quit=EXCEL, opt_exit=False)
logger.info("Closed Excel applications.")

######################## ↑ EXCEL DATA ↑ ########################     

######################## ↓ OUTLOOK ↓ ########################
#Instance of Outlook
OUTLOOK = client.Dispatch('Outlook.Application')
logger.info("Opened Outlook application.")

#Create a message
MESSAGE = OUTLOOK.CreateItem(0)

#Format Send To
SEND_TO_LIST = []
for LINE in USER_SEND_TO_LIST:
    try: #try to format 'recipient' with 'value' found in Excel
        SEND_TO_LIST.append(LINE.format_map(SafeDict(VAL_DICT))) #format by key
    except (KeyError, ValueError): #a value was not found for a key
        SEND_TO_LIST.append(LINE) #mantaining original line
        logger.warning("Key/Value error formating '{}' user 'sent to' settings !".format(LINE))
        pass
    except:
        logger.exception("Error formating '{}' user 'send to' settings !!!".format(LINE))
        pass

#Format Subject
SUBJECT=''
try: #try to format 'subject' with 'value' found in Excel
    SUBJECT = USER_SUBJECT.format_map(SafeDict(VAL_DICT)) #format by key
except (KeyError, ValueError): #a value was not found for a key
    SUBJECT = USER_SUBJECT #mantaining original line
    logger.warning("Key/Value error formating '{}' user 'subject' settings !".format(USER_SUBJECT))
    pass
except:
    logger.exception("Error formating '{}' user 'subject' settings !!!".format(USER_SUBJECT))
    pass

#List Image files
try: 
    LIST_DIR = os.listdir(FILES_DIR)
except:
    logger.exception("Error listing files in folder '{}' !!! ".format(FILES_DIR))
    safe_exit([*WORKBOOK, MESSAGE], [EXCEL, OUTLOOK], [file_handler, stream_handler], APP_ROOT)
LIST_IMG = [IMG for IMG in LIST_DIR if IMG.lower().endswith(('.png', '.jpg', '.jpeg', '.gif'))] #filter list of files by image
if len(LIST_IMG) == 0: #no image in folder
    logger.warning("No image in files folder '{}' !".format(FILES_DIR))

#Set properties to Image files
for IMAGE in LIST_IMG:
    try: #set microsoft properties to use CID to insert image in HTML
        attachment = MESSAGE.Attachments.Add(FILES_DIR + '/' + IMAGE)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", IMAGE.split(sep='.')[0])
        logger.info("Properties set to image '{}'.".format(IMAGE))
        #To add <img src="cid:IMAGE"> inside HTML file
    except:
        logger.exception("Error setting properties to image '{}' !!!".format(IMAGE))
        pass

#Format HTML Body
HTML_LIST = get_list_from_txt_file(HTML_FILE_PATH) #Divide HTML by lines
for INDEX, LINE in enumerate(HTML_LIST):
    try: #format line by line to not be stoped by errors
        HTML_LIST[INDEX] = LINE.format_map(SafeDict(VAL_DICT)) #format by key
    except (KeyError, ValueError): #error caused by false dict key '{' read
        HTML_LIST[INDEX] = LINE #mantaining original line
        logger.debug("False Key got in line '{}' in text '{}' ###".format(INDEX, LINE))
        pass
    except Exception:
        logger.exception("Error formating line '{}' !!!".format(LINE))
        safe_exit([*WORKBOOK, MESSAGE], [EXCEL, OUTLOOK], [file_handler, stream_handler], APP_ROOT)
HTML_BODY = '\n'.join(HTML_LIST) #join separated lines in one string

#Set the message properties
MESSAGE.To = ';'.join(SEND_TO_LIST) #str
MESSAGE.Subject = SUBJECT #str
MESSAGE.HTMLBody = HTML_BODY #str

#Ask for attachment files to email
DIALOG_TITLE = "Select one file to attach to the email or Cancel"
DIALOG_2ND_TITLE = "Select another file to attach to the email or Cancel"
ATT_FILE_LIST = ask_for_files(APP_ROOT, DIALOG_TITLE, 
                              at_least_one=False, more_than_one=True, second_title=DIALOG_2ND_TITLE)

#Attach files to email
for ATT_FILE_PATH in ATT_FILE_LIST:
    try:
        MESSAGE.Attachments.Add(ATT_FILE_PATH) #Attach
        logger.info("Attached file '{}' to the email".format(ATT_FILE_PATH))
    except:
        logger.exception("Error attaching file '{}' to the email".format(ATT_FILE_PATH))

#Display the message to user review
MESSAGE.Display()

#Send the message
#MESSAGE.Send()
######################## ↑ OUTLOOK ↑ ######################## 

safe_exit(to_Close=WORKBOOK, to_Quit=EXCEL, handler=[file_handler, stream_handler], app_root=APP_ROOT)