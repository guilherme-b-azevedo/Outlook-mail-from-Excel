"""
Docstring for the utils_general.py module.

This module contains general functions that can be useful in diferent
situations.
Contains Classes:
    SafeDict: Return the key not found between brackets.
Contains Functions:
    get_index_by_str: Find an index requested by a string.
"""
import logging

# Get Logger
if __name__ != '__main__':  # if this file was imported
    logger = logging.getLogger("__main__")


# Class definition
class SafeDict(dict):
    """A dictionary that returns '{key}' when the key that do not exists.

    This class behaves exactly like a dictionary except for its behavior when
    trying to access a key that does not exist in the dictionary. In this case,
    instead of an error being raised, the key itself is returned between braces
    It is commonly used in cases that is necessary a dictionary that do not
    raise a KeyError when a key that do not exists is passed. Example of
    usage: 'str.format_map(SafeDict(dict))'.
    """

    def __missing__(self, key):
        """Return the key not found between brackets.

        This function returns the missing key between brackets, like '{key}'
        when a key that does not exists in the dictionary is tried to be
        accessed. In a normal dictionary (dict type) in this cases an error
        is raised.

        Parameters
        ----------
        key : str
            A key that matches with the value required, if it exists.

        Returns
        -------
        str
            A string containing the key itself between braces like '{key}'.
        """
        return '{' + key + '}'


# Functions defition
def get_index_by_str(idx_text, opt_list=[], minus_one=True):
    """Find an index requested by a string.

    This function get an index requested by a string, wich can be only numeric
    or not. In the case of a numeric string, a convertion from str to int is
    done and returned, with the same value or 'minus_one'. In the case of a str
    not numeric, the index will be found by searching for the string in a list
    of options ('opt_list') and returning the integer index where the string
    was found. In the case of the string ('idx_text') passed is not numeric
    and not found inside the 'opt_list', the function returns None.

    Parameters
    ----------
    idx_text : str
        A text that should be identified on a list or a numeric string to be
        converter to integer.
    opt_list : Iterable, optional
        A list of strings that you want to find the correct index based on a
        string passed.
        The default is [].
    less_one : boolean, optional
        Used only in case of a numeric string passed, this returns the same
        value of the str converted to int when set to False. When set to True,
        it returns the value minus 1.
        The default is True.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    if idx_text.isnumeric():  # only integer number in the string
        if (
                len(opt_list) == 0 or
                (len(opt_list) > 0 and len(opt_list) >= int(idx_text))):
            if minus_one:
                return int(idx_text) - 1
            else:
                return int(idx_text)
    else:
        for idx, option in enumerate(opt_list):
            if idx_text in option:  # if option contains the idx_text suggested
                return idx
    logger.warning("Index not found for user defined index '{}' inside options"
                   " '{}' !".format(idx_text, opt_list))
    return None  # in case of not finding a valid index
