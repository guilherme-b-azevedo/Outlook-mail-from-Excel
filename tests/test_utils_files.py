"""
Docstring for the test_utils_files.py module.

This module contains tests of the module utils_files.py.

To execute 'tests' folder, from the Prompt, cd to the root folder (top) and run
python -m pytest
"""
import pytest
from utils import utils_files as u_fil


def test_func_safe_exit_1():
    assert u_fil.safe_exit("not", "correct", "object", "type",
                           False) is None, (
        "Failed case: incorrect types of objects to close.")


def test_func_safe_exit_2():
    assert u_fil.safe_exit([], [], [], [], False) is None, (
        "Failed case: empty list of objects to close.")


def test_func_safe_exit_3():
    with pytest.raises(SystemExit):
        u_fil.safe_exit(opt_exit=True)


@pytest.fixture
def fixt_txt_file_path():  # create a txt file in the test folder
    with open('/'.join(__file__[0:-3].split('\\')) + ".txt", "w") as file:
        file.write("1" + "\n" + "2" + "\n" + "{C}")
    return '/'.join(__file__[0:-3].split('\\')) + ".txt"


@pytest.mark.parametrize("fixt_txt_file_path,"
                         "remove_header,"
                         "headers,"
                         "dict_f,"
                         "raise_nf,"
                         "res",
                         [(fixt_txt_file_path,
                           False, [], {}, True, ['1', '2', '{C}']),
                          (fixt_txt_file_path,
                           True, [], {}, True, ['2', '{C}']),
                          (fixt_txt_file_path,
                           True, ['No'], {'D': '4'}, True, ['1', '2', '{C}']),
                          (fixt_txt_file_path,
                           True, ['1'], {'C': '3'}, True, ['2', '3'])],
                         indirect=["fixt_txt_file_path"])
def test_get_list_from_txt_file_by_line_1(fixt_txt_file_path, remove_header,
                                          headers, dict_f, raise_nf, res):
    assert u_fil.get_list_from_txt_file_by_line(fixt_txt_file_path,
                                                remove_header,
                                                headers, dict_f,
                                                raise_nf) == res


def test_delete_file_1(fixt_txt_file_path):
    assert u_fil.delete_file(fixt_txt_file_path) is True


def test_delete_file_2():
    assert u_fil.delete_file('/'.join(__file__[0:-3].split('\\')) +
                             ".jpg") is False


def test_get_list_from_txt_file_by_line_2():
    with pytest.raises(FileNotFoundError):
        u_fil.get_list_from_txt_file_by_line(
            '/'.join(__file__[0:-3].split('\\')) + ".txt",
            raise_not_found=True)


def test_get_list_from_txt_file_by_line_3():
    assert u_fil.get_list_from_txt_file_by_line(
        '/'.join(__file__[0:-3].split('\\')) + ".txt",
        raise_not_found=False) == []
