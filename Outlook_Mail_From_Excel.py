"""Created on Fri Oct 30 10:07:23 2020.

@author: Guilherme Bresciani de Azevedo
"""
# TO DO
# Try to format value got from excel based on 'CELL.NumberFormat' info.

import os
import time
import sys
import tkinter as tk
from tkinter import filedialog
import logging
from PIL import ImageGrab
import win32com.client as client
from utils import utils_general as u_gen
from utils import utils_files as u_fil
from utils import utils_excel as u_exc

# Global declaration
TIME_TO_SLEEP = 5


# Window app declaration
APP_ROOT = tk.Tk()
APP_ROOT.attributes("-topmost", True)  # to open dialogs in front of all
#APP_ROOT.lift()
APP_ROOT.withdraw()  # hide application main window
try:
    APP_ROOT.iconbitmap(os.getcwd() + '\\icon.ico')
except Exception:
    pass


# Get email Template .htm* file
HTML_FILE_PATH = filedialog.askopenfilename(parent=APP_ROOT,
                                            title="Select email template",
                                            filetypes=[('HTML files',
                                                       ['*.htm', '*.html'])])
try:  # try to read html file on path informed by user
    with open(HTML_FILE_PATH) as USER_FILE:
        HTML_BODY = USER_FILE.read()  # read as one string
except Exception:
    print("Error reading HTML file '{}' !!!".format(HTML_FILE_PATH))
    time.sleep(TIME_TO_SLEEP)
    sys.exit()
if len(HTML_BODY) == 0:
    print("File empty '{}' !!!".format(HTML_FILE_PATH))
    time.sleep(TIME_TO_SLEEP)
    sys.exit()


# Define files folder
LIST_POSSIBLE_FOLDERS = ['_arquivos', '_files', '_file', '_fichiers']
FILES_DIR = ''
for POSSIBLE_FOLDER in LIST_POSSIBLE_FOLDERS:
    # check if files folder exists in same path o HTML file
    if os.path.exists(HTML_FILE_PATH.split('.htm')[0] + POSSIBLE_FOLDER):
        FILES_DIR = HTML_FILE_PATH.split('.htm')[0] + POSSIBLE_FOLDER
        break
if FILES_DIR == '':  # folder don't exists with expected names in the folder
    print("No file folder corresponding to file '{}' was found in path '{}'"
          " !!!".format('/'.join(HTML_FILE_PATH.split('/')[0:-1]),
                        HTML_FILE_PATH.split('/')[-1]))
    time.sleep(TIME_TO_SLEEP)
    sys.exit()


# Log declaration
u_fil.delete_file(FILES_DIR + '/' + 'log.log')  # logging only last time run

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


# Check Manual values setting file .txt
TXT_FILE_PATH = FILES_DIR + "/setting_value_manually.txt"
USER_VAL_MAN_LIST = u_fil.get_list_from_txt_file_by_line(TXT_FILE_PATH,
                                                         remove_header=True,
                                                         raise_not_found=False)

# Start dict of values to replace
VAL_DICT = {}
for ID in USER_VAL_MAN_LIST:  # IDs defined in setting_value_manually.txt
    USER_MAN_VALUE = str(input("Write a value for the ID tag named '{}':\n"
                               .format(ID)))  # Ask for value of a key
    VAL_DICT[ID] = USER_MAN_VALUE

# Check Image setting file .txt
TXT_FILE_PATH = FILES_DIR + "/setting_image.txt"
USER_IMG_LIST = u_fil.get_list_from_txt_file_by_line(TXT_FILE_PATH, True, [],
                                                     u_gen.SafeDict(VAL_DICT))

# Check Value setting file .txt
TXT_FILE_PATH = FILES_DIR + "/setting_value.txt"
USER_VAL_LIST = u_fil.get_list_from_txt_file_by_line(TXT_FILE_PATH, True, [],
                                                     u_gen.SafeDict(VAL_DICT))

# Check Send To setting file.txt
TXT_FILE_PATH = FILES_DIR + "/setting_send_to.txt"
USER_SEND_TO_LIST = u_fil.get_list_from_txt_file_by_line(
                            TXT_FILE_PATH, True, [], u_gen.SafeDict(VAL_DICT))

# Check Subject setting file.txt
TXT_FILE_PATH = FILES_DIR + "/setting_subject.txt"
USER_SUBJECT_LIST = u_fil.get_list_from_txt_file_by_line(
                            TXT_FILE_PATH, True, [], u_gen.SafeDict(VAL_DICT))
USER_SUBJECT = USER_SUBJECT_LIST[0]


# Ask for Excel files to get information from
DIALOG_TITLE = "Select one Excel file to get information"
DIALOG_2ND_TITLE = "Select another Excel file to get information or Cancel"
DIALOG_FILE_TYPES = [('Excel files', ['*.xlsx', '*.xlsx', '*.xlsm'])]
XLS_FILE_LIST = u_fil.ask_for_files(APP_ROOT, DIALOG_TITLE,
                                    DIALOG_2ND_TITLE, DIALOG_FILE_TYPES,
                                    more_than_one=True)

# Check if at least one Excel file was selected
if XLS_FILE_LIST is None:  # user closed the first ask dialog
    u_fil.safe_exit(handler=[file_handler, stream_handler],
                    app_root=APP_ROOT)  # exit and log intentional exit

# List names of Worsheets for indexing data collection by name
XLS_FILE_NAMES = [FILE_PATH.split('/')[-1] for FILE_PATH in XLS_FILE_LIST]

################################# ↓ EXCEL ↓ #################################
# New instance of Excel
EXCEL = client.DispatchEx('Excel.Application')
logger.info("Opened Excel application.")
if int(EXCEL.Version[0:2]) < 16:
    #EXCEL.ScreenUpdating = True
    #EXCEL.DisplayAlerts = True
    EXCEL.Visible = True  # necessary to .CopyPicture() in Version 15.0 (2013)

# Open workbooks
WORKBOOK = []
for PATH in XLS_FILE_LIST:
    try:
        WORKBOOK.append(EXCEL.Workbooks.Open(PATH))
        logger.info("Opened Workbook '{}'.".format(PATH))
    except Exception:
        logger.error("Error opening the Workbook '{}' !!!".format(PATH))
        u_fil.safe_exit(to_Close=WORKBOOK, to_Quit=EXCEL,
                        handler=[file_handler, stream_handler],
                        app_root=APP_ROOT)

# List worksheets in same index order of WORKBOOK list
SHEET_LISTS = []
for BOOK in WORKBOOK:
    SHEET_LIST = []
    for SHEET in BOOK.Worksheets:
        SHEET_LIST.append(SHEET.Name)
    SHEET_LISTS.append(SHEET_LIST)
logger.info("Listed Worsheets by Workbook '{}'.".format(SHEET_LISTS))

######################## ↓ GET VALUES ↓ ########################
for SETUP in USER_VAL_LIST:  # SETUP = "ID	Sheet	Cell"
    # format setup with values previously got manually and automatically
    SETUP = SETUP.format_map(u_gen.SafeDict(VAL_DICT))
    KEY = str(SETUP.split(sep='\t')[0])  # dict key to format_map
    BOOK_INDEX = u_gen.get_index_by_str(SETUP.split(sep='\t')[1],
                                        XLS_FILE_NAMES)
    if BOOK_INDEX is None:
        logger.warning("Could not found Workbook for setup '{}' !"
                       .format(SETUP))
        next
    BOOK = WORKBOOK[BOOK_INDEX]  # user Book=1 will access by index 0
    SHEET_INDEX = u_gen.get_index_by_str(SETUP.split(sep='\t')[2],
                                         SHEET_LISTS[BOOK_INDEX])
    if SHEET_INDEX is None:
        logger.warning("Could not found Worksheet for setup '{}' !"
                       .format(SETUP))
        next
    SHEET = BOOK.Sheets[SHEET_INDEX]  # user Sheet=1 will access by index 0
    LOCATION = SETUP.split(sep='\t')[3]
    CELL = u_exc.get_cell(SHEET, LOCATION)  # decode LOCATION if logic present
    if CELL is not None:  # successfully got cell
        try:  # try to get .Value attribute of cell
            VAL_DICT[KEY] = str(CELL.Value)
        except Exception:
            logger.warning("Value not found for Cell in setup '{}' !"
                           .format(SETUP))
    else:
        logger.warning("Cell not found for Value setup '{}' !".format(SETUP))
# Log and print values found to user
logger.info("Values found for replacement are '{}'.".format(VAL_DICT))
print("\n")
print("Values found for replacement are:")
[print(KEY + ": " + VAL_DICT[KEY]) for KEY in VAL_DICT]
print("\n")
######################## ↑ GET VALUES ↑ ########################

######################## ↓ GET IMAGES ↓ ########################
for SETUP in USER_IMG_LIST:  # SETUP = "ID	Sheet	Type	Cell/Num"
    # format setup with values previously got manually and automatically
    SETUP = SETUP.format_map(u_gen.SafeDict(VAL_DICT))
    IMAGE_ID = SETUP.split(sep='\t')[0]  # image file with extension
    BOOK_INDEX = u_gen.get_index_by_str(SETUP.split(sep='\t')[1],
                                        XLS_FILE_NAMES)
    if BOOK_INDEX is None:
        logger.warning("Could not found Workbook for setup '{}' !"
                       .format(SETUP))
        u_fil.delete_file(FILES_DIR + '/' + IMAGE_ID)  # delete wrong old image
        next
    BOOK = WORKBOOK[BOOK_INDEX]  # user Book=1 will access by index 0
    SHEET_INDEX = u_gen.get_index_by_str(SETUP.split(sep='\t')[2],
                                         SHEET_LISTS[BOOK_INDEX])
    if SHEET_INDEX is None:
        logger.warning("Could not found Worksheet for setup '{}' !"
                       .format(SETUP))
        u_fil.delete_file(FILES_DIR + '/' + IMAGE_ID)  # delete wrong old image
        next
    SHEET = BOOK.Sheets[SHEET_INDEX]  # user Sheet=1 will access by index 0
    TYPE = SETUP.split(sep='\t')[3]  # Type = 'table' or 'chart'
    LOCATION = SETUP.split(sep='\t')[4]
    if TYPE == 'table':
        BGN_LOCATION = str(LOCATION.split(sep=':')[0])  # get A1 of range A1:B2
        END_LOCATION = str(LOCATION.split(sep=':')[1])  # get B2 of range A1:B2
        BGN_CELL = u_exc.get_cell(SHEET, BGN_LOCATION)  # decode logic present
        END_CELL = u_exc.get_cell(SHEET, END_LOCATION)  # decode logic present
        if (BGN_CELL is not None and
                END_CELL is not None):  # location could be recognized
            try:  # try to select and copy range
                RANGE = SHEET.Range(BGN_CELL, END_CELL)
                RANGE.CopyPicture(Appearance=1, Format=2)
                try:  # try to save image copied
                    ImageGrab.grabclipboard().save(FILES_DIR + '/' + IMAGE_ID)
                except Exception:
                    logger.warning("Could not save image for setup '{}' !"
                                   .format(SETUP))
                    u_fil.delete_file(FILES_DIR + '/' + IMAGE_ID)  # delete old
                    next
            except Exception:
                logger.warning("Range not found for setup '{}' !"
                               .format(SETUP))
                u_fil.delete_file(FILES_DIR + '/' + IMAGE_ID)  # delete old
                next
        else:
            logger.warning("Cell not found for Value setup '{}' !"
                           .format(SETUP))
            u_fil.delete_file(FILES_DIR + '/' + IMAGE_ID)  # delete old
            next
    elif TYPE == 'chart':
        try:
            CHART = SHEET.ChartObjects(int(LOCATION))  # select chart by number
            CHART.Activate()  # avoid exporting an image with nothing inside
            try:
                CHART.Chart.Export(FILES_DIR + '/' + IMAGE_ID)
            except Exception:
                logger.warning("Could not save image for setup '{}' !"
                               .format(SETUP))
                u_fil.delete_file(FILES_DIR + '/' + IMAGE_ID)  # delete old
                next
        except Exception:
            logger.warning("Chart no reachable for setup '{}' !".format(SETUP))
            u_fil.delete_file(FILES_DIR + '/' + IMAGE_ID)  # delete old
            next
    else:  # user defined type not recognizable
        logger.warning("Type not recognized for setup '{}' !".format(SETUP))
        u_fil.delete_file(FILES_DIR + '/' + IMAGE_ID)  # delete old
        next
# Image not reachable do not raise error, only warn and assure no wrong image
######################## ↑ GET IMAGES ↑ ########################

# Close Workbooks and Excel
u_fil.safe_exit(to_Close=WORKBOOK, to_Quit=EXCEL, opt_exit=False)
logger.info("Closed Excel applications.")

################################# ↑ EXCEL ↑ #################################

################################ ↓ OUTLOOK ↓ ################################
# Instance of Outlook
OUTLOOK = client.Dispatch('Outlook.Application')
logger.info("Opened Outlook application.")

# Create a message
MESSAGE = OUTLOOK.CreateItem(0)

# Format Send To
SEND_TO_LIST = []
for LINE in USER_SEND_TO_LIST:
    try:  # try to format 'recipient' with 'value' found in Excel
        SEND_TO_LIST.append(LINE.format_map(u_gen.SafeDict(VAL_DICT)))
    except (KeyError, ValueError):  # a value was not found for a key
        SEND_TO_LIST.append(LINE)  # mantaining original line
        logger.warning("Key/Value error formating '{}' user 'sent to' settings"
                       " !".format(LINE))
        pass
    except Exception:
        logger.exception("Error formating '{}' user 'send to' settings !!!"
                         .format(LINE))
        pass

# Format Subject
SUBJECT = ''
try:  # try to format 'subject' with 'value' found in Excel
    SUBJECT = USER_SUBJECT.format_map(u_gen.SafeDict(VAL_DICT))
except (KeyError, ValueError):  # a value was not found for a key
    SUBJECT = USER_SUBJECT  # mantaining original line
    logger.warning("Key/Value error formating '{}' user 'subject' settings"
                   " !".format(USER_SUBJECT))
    pass
except Exception:
    logger.exception("Error formating '{}' user 'subject' settings !!!"
                     .format(USER_SUBJECT))
    pass

# List files
try:
    LIST_DIR = os.listdir(FILES_DIR)
except Exception:
    logger.exception("Error listing files in folder '{}' !!! "
                     .format(FILES_DIR))
    u_fil.safe_exit([*WORKBOOK, MESSAGE], [EXCEL, OUTLOOK],
                    [file_handler, stream_handler], APP_ROOT)
# Filter list of files by image
LIST_IMG = [IMG for IMG in LIST_DIR
            if IMG.lower().endswith(('.png', '.jpg', '.jpeg', '.gif'))]
if len(LIST_IMG) == 0:  # no image in folder
    logger.warning("No image in files folder '{}' !".format(FILES_DIR))

# Set properties to Image files
for IMAGE in LIST_IMG:
    try:  # set microsoft properties to use CID to insert image in HTML
        attachment = MESSAGE.Attachments.Add(FILES_DIR + '/' + IMAGE)
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
            IMAGE.split(sep='.')[0])
        logger.info("Properties set to image '{}'.".format(IMAGE))
        # To add <img src="cid:IMAGE"> inside HTML file
    except Exception:
        logger.exception("Error setting properties to image '{}' !!!"
                         .format(IMAGE))
        pass

# Format HTML Body
HTML_LIST = u_fil.get_list_from_txt_file_by_line(HTML_FILE_PATH)  # List lines
for INDEX, LINE in enumerate(HTML_LIST):
    try:  # format line by line to not be stoped by errors
        HTML_LIST[INDEX] = LINE.format_map(u_gen.SafeDict(VAL_DICT))  # format
    except (KeyError, ValueError):  # error caused by false dict key '{' read
        HTML_LIST[INDEX] = LINE  # mantaining original line
        logger.debug("False Key got in line '{}' in text '{}' ###"
                     .format(INDEX, LINE))
        pass
    except Exception:
        logger.exception("Error formating line '{}' !!!".format(LINE))
        u_fil.safe_exit([*WORKBOOK, MESSAGE], [EXCEL, OUTLOOK],
                        [file_handler, stream_handler], APP_ROOT)
HTML_BODY = '\n'.join(HTML_LIST)  # join separated lines in one string

# Set the message properties
MESSAGE.To = ';'.join(SEND_TO_LIST)  # str
MESSAGE.Subject = SUBJECT  # str
MESSAGE.HTMLBody = HTML_BODY  # str

# Ask for attachment files to email
DIALOG_TITLE = "Select one file to attach to the email or Cancel"
DIALOG_2ND_TITLE = "Select another file to attach to the email or Cancel"
ATT_FILE_LIST = u_fil.ask_for_files(APP_ROOT, DIALOG_TITLE, DIALOG_2ND_TITLE,
                                    at_least_one=False, more_than_one=True)

# Attach files to email
for ATT_FILE_PATH in ATT_FILE_LIST:
    try:
        MESSAGE.Attachments.Add(ATT_FILE_PATH)  # Attach
        logger.info("Attached file '{}' to the email".format(ATT_FILE_PATH))
    except Exception:
        logger.exception("Error attaching file '{}' to the email"
                         .format(ATT_FILE_PATH))

# Display the message to user review
MESSAGE.Display()

# Send the message
#MESSAGE.Send()
################################ ↑ OUTLOOK ↑ ################################ 

u_fil.safe_exit(to_Close=WORKBOOK, to_Quit=EXCEL,
                handler=[file_handler, stream_handler], app_root=APP_ROOT)
