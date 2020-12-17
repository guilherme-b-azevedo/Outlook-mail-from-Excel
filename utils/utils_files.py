"""
Docstring for the utils_files.py module.

This module contains general functions that can be useful for handling files.
Contains Functions:
    safe_exit: Close files, handlers, tk application and exit interpreter.
    ask_for_files: Open a file dialog asking user for one or more files.
    get_list_from_txt_file_by_line: Read a text file and make a list of its
    lines.
    delete_file: Delete a file if it exists.
"""
from collections.abc import Iterable
import os
import sys
import logging
from tkinter import filedialog

# Get Logger
if __name__ != '__main__':  # if this file was imported
    logger = logging.getLogger("__main__")


# Functions defition
def safe_exit(to_Close=None, to_Quit=None,
              handler=None, app_root=None,
              opt_exit=True):
    r"""Close files, handlers, tk application and exit interpreter.

    Function normally used to terminate the software execution correctly,
    by closing files first and terminating after.

    This function close:
        Files: by calling methods .Close() and .Quit() to close files.
        Usually files opened by win32com lib like Excel files, closes this way.
        Log handlers: by calling methods .flush(), .close()
        and removeHandler().
        Tk application root: by calling method .destroy().
        Interpreter itself: by calling sys.exit().

    Parameters
    ----------
    to_Close : object or iterable of objects, optional
        Objects that uses method .Close().
    to_Quit : object or iterable of objects, optional
        Objects that uses method .Quit().
    handler : handler of logger, optional
        Handler created by logging.FileHandler() or logging.StreamHandler().
    app_root : tk application, optional
        Tk application created by APP_ROOT = tk.Tk().
    opt_exit : boolean, default=True
        Defines if the software will be closed (True) or not (False).

    Raises
    ------
    No chance of raising.

    Returns
    -------
    None
    """
    try:
        logger.info("Closing...")
    except Exception:
        pass
    if isinstance(to_Close, Iterable):
        for item in to_Close:
            try:
                item.Close()
            except Exception:
                pass
    elif to_Close is not None:
        try:
            to_Close.Close()
        except Exception:
            pass
    if isinstance(to_Quit, Iterable):
        for item in to_Quit:
            try:
                item.Quit()
            except Exception:
                pass
    elif to_Quit is not None:
        try:
            to_Quit.Quit()
        except Exception:
            pass
    if isinstance(handler, Iterable):
        for hnd in handler:
            try:
                hnd.flush()
                hnd.close()
                logger.removeHandler(hnd)
            except Exception:
                pass
    elif handler is not None:
        try:
            handler.flush()
            handler.close()
            logger.removeHandler(handler)
        except Exception:
            pass
    if app_root is not None:
        try:
            app_root.destroy()
        except Exception:
            pass
    if opt_exit:
        sys.exit()


def ask_for_files(dialog_parent, dialog_title, second_title='',
                  file_types=[('All files', '*.*')],
                  at_least_one=True, more_than_one=False):
    r"""Open a file dialog asking user for one or more files.

    The function receives at least the 'dialog_parent'
    (like 'APP_ROOT = tk.Tk()') and a 'dialog_title'
    (like 'Select a file dear user') to open a dialog asking the user to
    select some file, wich will return a full path for the file.
    Optionally, more than one file can be selected, by asking again an again
    until the user closes the last dialog opened, indicating that enough
    files was selected.
    Optionally, the selection a one file can be set as optional, by giving the
    user the choice to select some file or not.

    Parameters
    ----------
    dialog_parent : tk application root
        A tkinter application created by somethin like tk.Tk().
    dialog_title : str
        Objects that uses method .Quit().
    handler : handler of logger, optional
        Handler created by logging.FileHandler() or logging.StreamHandler().
    app_root : tk application, optional
        Tk application created by APP_ROOT = tk.Tk().
    opt_exit : boolean, default=True
        Defines if the software will be closed (True) or not (False).

    Raises
    ------
    None

    Returns
    -------
    None
        When no file is selected (by dialog is closing or canceling) and
        'at_least_one' file was required.
    PATH : str
        Full path of the file selected when only one file selection is allowed,
        (more_than_one=False)
    PATH_LIST : list
        List containing a full path of the diles selected by the user, with '/'
        separating the folders. It can be an empty list if a selection of file
        is optinal (at_least_one=False).
    """
    # Copy the title of the first asking dialog to the second if not passed
    if second_title == '':
        second_title = dialog_title

    PATH_LIST = []
    PATH = "first time asking"
    while len(PATH) > 0:
        PATH = filedialog.askopenfilename(parent=dialog_parent,
                                          title=dialog_title,
                                          filetypes=file_types)
        if len(PATH) > 0:  # user selected a file
            logger.info("User selected file path '{}'".format(PATH))
            if more_than_one:  # one more dialog will be opened asking for file
                PATH_LIST.append(PATH)
                dialog_title = second_title
                next
            else:  # 1 file selected and only 1 required
                return PATH
        elif len(PATH_LIST) > 0:  # files selected and more_than_one=True
            return PATH_LIST
        elif at_least_one:  # 1 file is required and user did not selected
            logger.error("A file was not selected when asked '{}' !!!"
                         .format(dialog_title))
            return None
        else:
            logger.info("A file was not selected when asked '{}'."
                        .format(dialog_title))
            return PATH_LIST


def get_list_from_txt_file_by_line(file_path,
                                   remove_header=False, list_headers=[],
                                   dict_format={}, raise_not_found=True):
    """Read a text file and make a list of its lines.

    The function receives a text 'file_path' and separate and writen line in a
    string, and compose a list with this string. It removes blank lines and
    white spaces at the end of the lines. An empty file do raise any error,
    only returns an empty list.
    Optionally, it removes the first line as a header (remove_header=True).
    Optionally, it verifies if the first line is a header to be removed, by
    looking to the 'list_headers' if the first line matches.
    Optionally, it can format lines using a dictionary of terms to replace.
    Optionally, it can raise an error if the 'file_path' was not found.

    Parameters
    ----------
    file_path : str
        Full path of the text file to read and get a list by lines.
    remove_header : boolean, optional
        Choice to remove the first line of the list of lines.
        The default is False.
    list_headers : list of str, optional
        List of possibilities of hearders to be presente in the first line,
        that should be removed when 'remove_header=True'.
        The default is [].
    dict_format : dict of str, optional
        Dictionary of keys to be replaced by values line by line.
        The dictionary looks like {'key':'value'} and the 'value' will be
        replaced by the key when '{key}' is found in the some line. Watch out
        the KeyError if the key is not found in the dictionary.
        The default is {}.
    raise_not_found : boolean, optional
        Option to raise (True) or not (False) the exception FileNotFoundError,
        when 'file_path' is not found.
        The default is True.

    Raises
    ------
    FileNotFoundError
        Raised when 'file_path' is not found only if raise_not_found=True.

    Returns
    -------
    USER_LIST : list of str
        A list containing every writen line of the text file,
        with or without the first line, formated or not by a dictionary.

    """
    USER_LIST = []
    # Open the txt file and list line by line removing white space at the end
    try:
        with open(file_path) as FILE:
            USER_LIST = [line.strip() for line in FILE if line.strip()]
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
    # Remove header from user list
    if remove_header and len(USER_LIST) > 0:
        if len(list_headers) == 0:  # if no possible headers was passed
            USER_LIST.pop(0)
        else:
            for TITLE in list_headers:
                if (
                        USER_LIST[0].casefold() == TITLE.casefold() or
                        USER_LIST[0].casefold() == TITLE.casefold() + 's'):
                    USER_LIST.pop(0)
                    logger.info("Header '{}' removed from file '{}'"
                                .format(USER_LIST[0], file_path))
    # Check exitence of information
    if len(USER_LIST) == 0:
        logger.warning("File empty '{}' !".format(file_path))
    else:
        logger.info("File read in '{}'.".format(file_path))
        if len(dict_format) > 0:  # Format lines using dictionary
            for INDEX, ITEM in enumerate(USER_LIST):
                try:
                    USER_LIST[INDEX] = ITEM.format_map(dict_format)
                except Exception:
                    logger.warning("Error formating setup '{}' in file '{}' "
                                   "using manually setted values !!!"
                                   .format(ITEM, file_path))
                    pass
    return USER_LIST


def delete_file(file_path):
    """Delete a file if it exists.

    This function checks for file existence  before to delete it. If the file
    do not exists, no error is raised. The function returns True if a file was
    deleted and False if the file do not exists

    Parameters
    ----------
    file_path : str
        Full path to the file that should be deleted.

    Returns
    -------
    bool
        True if the file was deleted. False if the file do not exists.

    """
    if os.path.exists(file_path):  # file exists
        os.remove(file_path)  # delete file
        return True
    else:  # file do not exists
        return False
