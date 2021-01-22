"""
Docstring for the utils_excel.py module.

This module contains general functions that can be useful to deal with excel.
Contains Functions:
    find_column: Find column based on search code.
    find_row: Find row based on search code.
    get_cell: Decode cell location
"""
import logging
import re

# Get Logger
if __name__ != '__main__':  # if this file was imported
    logger = logging.getLogger("__main__")


def find_column(sheet, coded_location):
    """Find column based on search code.

    This function receives an excel 'sheet' opened by win32com library and a
    str ('coded_location') with parameters that compose a search formula.
    This function will decode the str, understand the parameters passed and
    search for a cell that matches the 'coded_location' in the 'sheet'. If a
    cell is found, the function returns the column index, if not found will
    return '-1'.

    The 'coded_location' is composed by 5 parameters, separated by semicolon,
    all inside parentheses, like '(param1;param2;param3;param4;param5)'. In a
    nutshell, the search code for a column determines a line to be traversed
    until a comparison with a given value is successful, and a column offset
    can be applied after all.

    Detailing the parameters of the search code:
    Parameter 1 (col_in_row) – Row to be traversed.
    a.	When looking for a column, indicate a line to be parsed, like '2'.
    b.	Example in filling in the search formula:
        i.	(2;parameter2;parameter3;parameter4;parameter5)

    Parameter 2 (col_comp) – Comparison to be done.
    a.	The possibilities of comparison are:
        i.	“==” that means “equal than”.
        ii.	“!=” or “<>” that means “different than”.
    b.	Example in filling in the search formula:
        i.	(2;==;parameter3;parameter4;parameter5)

    Parameter 3 (col_text) – Cell value to be compared with.
    a.	The type of the comparison was set by the parameter 2,
        to be compared with whatever is written in parameter 3.
        So, the parameter 3 contains the complete content of the cell you want
        to check if it is equal or different, according to parameter 2.
    b.	Beware of cells that contain formatted values. It is common cases in
        which the value that appears inside the excel cell is 90% for example,
        but the actual value of the cell's content is 0.9, and the cell is only
        formatted so that it appears in percentage.
    c.	Keep in mind that you are looking for content that will make you sure
        that, from this column or row, you know exactly where is the cell you
        want to get information from.
    d.	Example in filling in the search formula:
        i.	(2;==;Project Name: ;parameter4;parameter5)

    Parameter 4 (col_offset) – Offset signal.
    a.	From a column/row that matched by de comparison defined by parameters
        2 and 3, you can take some column before or after that.
    b.	The possibilities of offset signal are:
        i.	“+” that takes a column after the cell matched by the comparison.
        ii.	“-” that takes a column before the cell matched by the comparison.
    c.	Example in filling in the search formula:
        i.	(2;==;Project Name: ;+;parameter5)

    Parameter 5 (col_off_value) – Offset quantity.
    a.	From a column/row that matched by de comparison defined by parameters
        2 and 3, the parameter 5 define how many columns/rows after or before,
        depending on parameter 4, the software will set the cell location.
    b.	The possibilities of offset quantity are any integer number.
        Example “1”, “2”, “3”.
    c.	Example in filling in the search formula:
        i.	(2;==;Project Name: ;+;1)

    Parameters
    ----------
    sheet : win32com object type
        A sheet opened by win32com lib like 'SHEET = BOOK.Sheets[SHEET_INDEX]'.
    coded_location : str
        A string that starts and ends with parentheses, with 5 parameters
        separeted by semicolon like '(1;==;any text;+;0)'.

    Returns
    -------
    SET_COL : int
        An integer number that represents the index of the column. It can be
        used by the function 'sheet.Cells(ROW, COLUMN)' of win32com lib.
        In case of not finding a column that matches with the search formula
        passed, it returns '-1'.
    """
    # coded_location example: (1;==;any text;+;0)
    SET_COL = -1
    # Parse columns +2 to go until the first empty cell
    for COLUMN in range(1, sheet.UsedRange.Columns.Count + 2):
        CELL_VALUE = sheet.Cells(
            int(coded_location.group('col_in_row')), COLUMN).Value
        if CELL_VALUE is None:  # match emp´ty cell with empty string
            CELL_VALUE = ""
        if coded_location.group('col_comp') == "==":
            if str(CELL_VALUE) == coded_location.group('col_text'):
                if coded_location.group('col_offset') == "+":
                    SET_COL = COLUMN+int(coded_location.group('col_off_value'))
                    break
                elif coded_location.group('col_offset') == "-":
                    SET_COL = COLUMN-int(coded_location.group('col_off_value'))
                    break
        elif (coded_location.group('col_comp') == "!=" or
              coded_location.group('col_comp') == "<>"):
            if str(CELL_VALUE) != coded_location.group('col_text'):
                if coded_location.group('col_offset') == "+":
                    SET_COL = COLUMN+int(coded_location.group('col_off_value'))
                    break
                elif coded_location.group('col_offset') == "-":
                    SET_COL = COLUMN-int(coded_location.group('col_off_value'))
                    break
        elif coded_location.group('col_comp') == ">=":
            if CELL_VALUE >= float(coded_location.group('col_text')):
                if coded_location.group('col_offset') == "+":
                    SET_COL = COLUMN+int(coded_location.group('col_off_value'))
                    break
                elif coded_location.group('col_offset') == "-":
                    SET_COL = COLUMN-int(coded_location.group('col_off_value'))
                    break
        elif coded_location.group('col_comp') == "<=":
            if CELL_VALUE <= float(coded_location.group('col_text')):
                if coded_location.group('col_offset') == "+":
                    SET_COL = COLUMN+int(coded_location.group('col_off_value'))
                    break
                elif coded_location.group('col_offset') == "-":
                    SET_COL = COLUMN-int(coded_location.group('col_off_value'))
                    break
    if SET_COL < 0:
        logger.warning("Column not found for logic '{}' !"
                       .format(coded_location.group('col')))
    else:
        logger.info("Column '{}' found for logic '{}'."
                    .format(SET_COL, coded_location.group('col')))
    return SET_COL


def find_row(sheet, coded_location):
    """Find row based on search code.

    This function receives an excel 'sheet' opened by win32com library and a
    str ('coded_location') with parameters that compose a search formula.
    This function will decode the str, understand the parameters passed and
    search for a cell that matches the 'coded_location' in the 'sheet'. If a
    cell is found, the function returns the row index, if not found will
    return '-1'.

    The 'coded_location' is composed by 5 parameters, separated by semicolon,
    all inside parentheses, like '(param1;param2;param3;param4;param5)'. In a
    nutshell, the search code for a row determines a line to be traversed
    until a comparison with a given value is successful, and a row offset
    can be applied after all.

    Detailing the parameters of the search code:
    Parameter 1 (lin_in_col) – Row to be traversed.
    a.	When looking for a row, indicate a line to be parsed, like 'B'.
    b.	Example in filling in the search formula:
        i.	(B;parameter2;parameter3;parameter4;parameter5)

    Parameter 2 (lin_comp) – Comparison to be done.
    a.	The possibilities of comparison are:
        i.	“==” that means “equal than”.
        ii.	“!=” or “<>” that means “different than”.
    b.	Example in filling in the search formula:
        i.	(B;!=;parameter3;parameter4;parameter5)

    Parameter 3 (lin_text) – Cell value to be compared with.
    a.	The type of the comparison was set by the parameter 2,
        to be compared with whatever is written in parameter 3.
        So, the parameter 3 contains the complete content of the cell you want
        to check if it is equal or different, according to parameter 2.
    b.	Beware of cells that contain formatted values. It is common cases in
        which the value that appears inside the excel cell is 90% for example,
        but the actual value of the cell's content is 0.9, and the cell is only
        formatted so that it appears in percentage.
    c.	Keep in mind that you are looking for content that will make you sure
        that, from this column or row, you know exactly where is the cell you
        want to get information from.
    d.	Example in filling in the search formula:
        i.	(B;!=;example of a cell content;parameter4;parameter5)

    Parameter 4 (lin_offset) – Offset signal.
    a.	From a column/row that matched by de comparison defined by parameters
        2 and 3, you can take some row before or after that.
    b.	The possibilities of offset signal are:
        i.	“+” that takes a row after the cell matched by the comparison.
        ii.	“-” that takes a row before the cell matched by the comparison.
    c.	Example in filling in the search formula:
        i.	(B;!=;example of a cell content;-;parameter5)

    Parameter 5 (lin_off_value) – Offset quantity.
    a.	From a column/row that matched by de comparison defined by parameters
        2 and 3, the parameter 5 define how many columns/rows after or before,
        depending on parameter 4, the software will set the cell location.
    b.	The possibilities of offset quantity are any integer number.
        Example “1”, “2”, “3”.
    c.	Example in filling in the search formula:
        i.	(B;!=;example of a cell content;-;5)

    Parameters
    ----------
    sheet : win32com object type
        A sheet opened by win32com lib like 'SHEET = BOOK.Sheets[SHEET_INDEX]'.
    coded_location : str
        A string that starts and ends with parentheses, with 5 parameters
        separeted by semicolon like '(A;==;any text;+;0)'.

    Returns
    -------
    SET_COL : int
        An integer number that represents the index of the row. It can be
        used by the function 'sheet.Cells(ROW, COLUMN)' of win32com lib.
        In case of not finding a row that matches with the search formula
        passed, it returns '-1'.
    """
    # coded_location example: (A;==;any text;+;0)
    SET_ROW = -1
    # Parse row +2 to go until the first empty cell
    for ROW in range(1, sheet.UsedRange.Rows.Count + 2):
        CELL_VALUE = sheet.Range(
            coded_location.group('lin_in_col') + str(ROW)).Value
        if CELL_VALUE is None:
            CELL_VALUE = ""  # match empty cell with empty string
        if coded_location.group('lin_comp') == "==":
            if str(CELL_VALUE) == coded_location.group('lin_text'):
                if coded_location.group('lin_offset') == "+":
                    SET_ROW = ROW + int(coded_location.group('lin_off_value'))
                    break
                elif coded_location.group('lin_offset') == "-":
                    SET_ROW = ROW - int(coded_location.group('lin_off_value'))
                    break
        elif (coded_location.group('lin_comp') == "!=" or
              coded_location.group('lin_comp') == "<>"):
            if str(CELL_VALUE) != coded_location.group('lin_text'):
                if coded_location.group('lin_offset') == "+":
                    SET_ROW = ROW + int(coded_location.group('lin_off_value'))
                    break
                elif coded_location.group('lin_offset') == "-":
                    SET_ROW = ROW - int(coded_location.group('lin_off_value'))
                    break
        elif coded_location.group('lin_comp') == ">=":
            if CELL_VALUE >= float(coded_location.group('lin_text')):
                if coded_location.group('lin_offset') == "+":
                    SET_ROW = ROW + int(coded_location.group('lin_off_value'))
                    break
                elif coded_location.group('lin_offset') == "-":
                    SET_ROW = ROW - int(coded_location.group('lin_off_value'))
                    break
        elif coded_location.group('lin_comp') == "<=":
            if CELL_VALUE <= float(coded_location.group('lin_text')):
                if coded_location.group('lin_offset') == "+":
                    SET_ROW = ROW + int(coded_location.group('lin_off_value'))
                    break
                elif coded_location.group('lin_offset') == "-":
                    SET_ROW = ROW - int(coded_location.group('lin_off_value'))
                    break
    if SET_ROW < 0:
        logger.warning("Row not found for logic '{}' !"
                       .format(coded_location.group('lin')))
    else:
        logger.info("Row '{}' found for logic '{}'."
                    .format(SET_ROW, coded_location.group('lin')))
    return SET_ROW


def get_cell(sheet, location):
    """Select a cell in an excel sheet based on user location string.

    This function receives an excel 'sheet' opened by win32com library and a
    standard string wich indicates the 'location' of a cell. This 'location'
    can be a letter, that indicates the column, followed by a integer number
    that indicates the row (like 'A1'). There is also the possibility to,
    instead of passing a letter or a number, pass a search formula, in cases
    that the exactly column or row is unkown and can be found by a comparison
    of the value of a cell. In any case the 'location' must recognizable by a
    PATTERN.

    Examples of standard possibilities of 'location':
    a.	Fixed column and fixed row:
        i.	'A1'
        ii.	'BR10'
    b.	Fixed column and search row:
        i.	'A(' + lin_in_col +';'+ lin_comp +';'+ lin_text +';'+ lin_offset +
            ';'+ lin_off_value +')'
        ii.	'BR(A;==;any text;+;0)'
    c.	Search column and fixed row:
        i.	'(' + col_in_row +';'+ col_comp +';'+ col_text +';'+ col_offset +
            ';'+ col_off_value +')1'
        ii.	'(1;==;any text;+;0)10'.
    d.	Search column and search row:
        i.	'(' + ';'.join([col_in_row, col_comp, col_text, col_offset,
            col_off_value]) + ')' + '(' + ';'.join([lin_in_col, lin_comp,
            lin_text, lin_offset, lin_off_value]) + ')'
        ii.	'(1;==;any text;+;0)(A;==;any text;+;0)'

    Meaning of the 5 parameters of the coded 'location':
    Parameter 1 (col_in_row/lin_in_col) – Column/Row to be traversed.
        a.	When looking for a column, indicate a line to be traversed, like
            '2'. When looking for a row, indicate a column to be traversed,
            such as ‘B’ for example.
        b.	Example in filling in the search formula:
            i.	(2;parameter2;parameter3;parameter4;parameter5)2
            ii.	B(B;parameter2;parameter3;parameter4;parameter5)

    Parameter 2 (col_comp/lin_comp) – Comparison to be done.
        a.	The possibilities of comparison are:
            i.	“==” that means “equal than”.
            ii.	“!=” or “<>” that means “different than”.
        b.	Example in filling in the search formula:
            i.	(2;==;parameter3;parameter4;parameter5)2
            ii.	B(B;!=;parameter3;parameter4;parameter5)

    Parameter 3 (col_text/lin_text) – Cell value to be compared with.
    a.	The type of the comparison was set by the parameter 2, to be compared
        with whatever is written in parameter 3. So, the parameter 3 contains
        the complete content of the cell you want to check if it is equal or
        different, according to parameter 2.
    b.	Beware of cells that contain formatted values. It is common cases in
        which the value that appears inside the excel cell is 90% for example,
        but the actual value of the cell's content is 0.9, and the cell is only
        formatted so that it appears in percentage.
    c.	Keep in mind that you are looking for content that will make you sure
        that, from this column or row, you know exactly where is the cell you
        need to get information from.
    d.	Example in filling in the search formula:
        i.	(2;==;Project Name: ;parameter4;parameter5)2
        ii.	B(B;!=;example of a cell content;parameter4;parameter5)

    4.	Parameter 4 (col_offset/lin_offset) – Offset signal.
    a.	From a column/row that matched by de comparison defined by parameters
        2 and 3, you can take some column/row before or after that.
    b.	The possibilities of offset signal are:
        i.	“+” that takes a column/row after the cell matched by the
        comparison.
        ii.	“-” that takes a column/row before the cell matched by the
        comparison.
    c.	Example in filling in the search formula:
        i.	(2;==;Project Name: ;+;parameter5)2
        ii.	B(B;!=;example of a cell content;-;parameter5)

    5.	Parameter 5 (col_off_value/lin_off_value) – Offset quantity.
    a.	From a column/row that matched by de comparison defined by parameters
        2 and 3, the parameter 5 define how many columns/rows after or before,
        depending on parameter 4, the software will find the cell location.
    b.	The possibilities of offset quantity are any integer number.
        Example “1”, “3”, “10”.
    c.	Example in filling in the search formula:
        i.	(2;==;Project Name: ;+;1)2
        ii.	B(B;!=;example of a cell content;-;5)

    Parameters
    ----------
    sheet : win32com object type
        A sheet opened by win32com lib like 'SHEET = BOOK.Sheets[SHEET_INDEX]'.
    location : str
        A string that indicates the location of a cell. It can be directly
        like "A1" or by a search formular that starts and ends with parentheses
        and contains 5 parameters separeted by semicolon like
        '(1;==;any text;+;0)(A;==;any text;+;0)'.

    Raises
    ------
    None

    Returns
    -------
    win32com object type
        A selected CELL of an excel file, which the value can be read using the
        attribute CELL.Value for example.
    None
        In case of not finding a valid cell based on the search formula passed.
    """
    # Set pattern for user setup of 'location'
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

    # Recognize pattern in user defined location
    location = re.match(PATTERN, str(location))  # same string but marked
    if location is None:  # location could not be marked as standard
        logger.warning("Standard not recognized for location setup '{}' !"
                       .format(location))
        return None

    # Decode location already recognized as standard
    if (location.group('col_in_row') is None and
            location.group('lin_in_col') is None):  # location == A1
        # return sheet.Range because location is str
        return sheet.Range(location.group('cell'))
    elif (location.group('col_in_row') is not None and
          location.group('lin_in_col') is None):  # location == (logic)1
        COLUMN = find_column(sheet, location)
        if COLUMN <= 0:  # if not found
            logger.warning("Column not found for setup '{}' !"
                           .format(location.group(0)))
            return None
        else:
            # return sheet.Cells because Column is int
            return sheet.Cells(int(location.group('lin')), COLUMN)
    elif (location.group('col_in_row') is None and
          location.group('lin_in_col') is not None):  # location == A(logic)
        ROW = find_row(sheet, location)
        if ROW <= 0:  # if not found
            logger.warning("Row not found for setup '{}' !"
                           .format(location.group(0)))
            return None
        else:
            return sheet.Range(location.group('col') + str(ROW))
    elif (location.group('col_in_row') is not None and
          location.group('lin_in_col') is not None):  # locatio==(logic)(logic)
        COLUMN = find_column(sheet, location)
        if COLUMN <= 0:  # if not found
            logger.warning("Column not found for setup '{}' !"
                           .format(location.group(0)))
            return None
        ROW = find_row(sheet, location)
        if ROW <= 0:  # if not found
            logger.warning("Row not found for setup '{}' !"
                           .format(location.group(0)))
            return None
        return sheet.Cells(ROW, COLUMN)
    # Cell not found do not raise error, only warn user and return None
