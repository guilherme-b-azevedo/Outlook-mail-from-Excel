"""
Docstring for the test_utils_excel.py module.

This module contains tests of the module utils_excel.py.

To execute 'tests' folder, from the Prompt, cd to the root folder (top) and run
python -m pytest
"""
import pytest
from utils import utils_excel as u_exc
import win32com.client as client


@pytest.fixture
def fixt_excel_file():  # create a excel file
    EXCEL = client.Dispatch('Excel.Application')  # open excel instance
    WB = EXCEL.Workbooks.Add()  # workbook creation
    WS = WB.Worksheets.Add()  # worksheet creation
    WS.Range("A1").Value = 1
    WS.Range("A2").Value = 2
    WS.Range("A1:A2").AutoFill(WS.Range("A1:A10"), 0)  # Type=0=xlFillDefault
    WS.Range("B1").Value = 10
    WS.Range("B2").Value = 20
    WS.Range("B1:B2").AutoFill(WS.Range("B1:B10"), 0)  # Type=0=xlFillDefault

    return EXCEL, WB, WS  # return app, workbook, worksheet with dummy values


@pytest.mark.parametrize("fixt_excel_file,"
                         "coded_location,"
                         "res",
                         [(fixt_excel_file, 'A1', 1),
                          (fixt_excel_file, '(1;==;1.0;+;0)1', 1),
                          (fixt_excel_file, '(1;==;10.0;-;1)1', 1),
                          (fixt_excel_file, '(1;==;;+;0)1', None),
                          (fixt_excel_file, '(1;!=;1.0;+;1)1', None),
                          (fixt_excel_file, '(1;!=;1.0;-;1)1', 1),
                          (fixt_excel_file, 'A(A;==;1.0;+;0)', 1),
                          (fixt_excel_file, 'A(A;==;2.0;-;1)', 1),
                          (fixt_excel_file, 'A(A;==;;+;0)', None),
                          (fixt_excel_file, 'A(A;!=;1.0;+;1)', 3),
                          (fixt_excel_file, 'A(A;!=;1.0;-;1)', 1),
                          (fixt_excel_file, '(1;==;1.0;+;0)(A;==;2.0;-;1)', 1),
                          (fixt_excel_file, '(1;>=;1.0;+;0)(A;>=;1.0;+;1)', 2),
                          (fixt_excel_file, '(1;>=;1.0;-;0)(A;>=;2.0;-;1)', 1),
                          (fixt_excel_file, '(1;<=;10;+;1)(A;<=;2;+;1)', 20),
                          (fixt_excel_file, '(1;<=;2;-;0)(A;<=;3;-;0)', 1)
                          ],
                         indirect=["fixt_excel_file"])
def test_get_cell_1(fixt_excel_file, coded_location, res):
    assert u_exc.get_cell(fixt_excel_file[2], coded_location).Value == res
    fixt_excel_file[1].Close(False)  # close workbook
    fixt_excel_file[0].Quit()  # close excel app


@pytest.mark.parametrize("fixt_excel_file,"
                         "coded_location,"
                         "res",
                         [(fixt_excel_file, '1A', None),
                          (fixt_excel_file, '(1;==;1.0;-;1)1', None),
                          (fixt_excel_file, 'A(A;==;1.0;-;1)', None),
                          (fixt_excel_file, '(1;==;1.0;-;1)(A;==;1.0;-;1)',
                           None),
                          (fixt_excel_file, '(1;==;1.0;-;0)(A;==;1.0;-;1)',
                           None),
                          (fixt_excel_file, '(1;==;1.0;-;1)(A;==;1.0;-;0)',
                           None),
                          ],
                         indirect=["fixt_excel_file"])
def test_get_cell_2(fixt_excel_file, coded_location, res):
    assert u_exc.get_cell(fixt_excel_file[2], coded_location) == res
    fixt_excel_file[1].Close(False)  # close workbook
    fixt_excel_file[0].Quit()  # close excel app
