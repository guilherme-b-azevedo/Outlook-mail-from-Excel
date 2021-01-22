"""
Docstring for the test_utils_general.py module.

This module contains tests of the module utils_general.py.

To execute 'tests' folder, from the Prompt, cd to the root folder (top) and run
python -m pytest
"""
# import pytest
from utils import utils_general as u_gen


def test_func_get_index_by_str_1():
    assert u_gen.get_index_by_str("0", [], False) == 0, (
        "Failed case: numeric str '0', no option list, minus one false.")


def test_func_get_index_by_str_2():
    assert u_gen.get_index_by_str("1", [], True) == 0, (
        "Failed case: numeric str '1', no option list, minus one true.")


def test_func_get_index_by_str_3():
    assert u_gen.get_index_by_str("0", [], True) is None, (
        "Failed case: numeric str '0', no option list, minus one true.")


def test_func_get_index_by_str_4():
    assert u_gen.get_index_by_str("3", [1, 2], True) is None, (
        "Failed case: numeric str '3', with option list len=2,"
        " minus one true.")


def test_func_get_index_by_str_5():
    assert u_gen.get_index_by_str("3", [7, 8, 9], True) == 2, (
        "Failed case: numeric str '3', with option list len=3,"
        " minus one true.")


def test_func_get_index_by_str_6():
    assert u_gen.get_index_by_str("txt", [], False) is None, (
        "Failed case: text str 'txt', no option list, minus one false.")


def test_func_get_index_by_str_7():
    assert u_gen.get_index_by_str("txt0", ["txt1", "txt2"], False) is None, (
        "Failed case: text str 'txt0', opt_list=['txt1', 'txt2'],"
        " minus one false.")


def test_func_get_index_by_str_8():
    assert u_gen.get_index_by_str("txt2", ["txt1", "txt2"], False) == 1, (
        "Failed case: text str 'txt2', opt_list=['txt1', 'txt2'],"
        " minus one false.")


def test_func_get_index_by_str_9():
    assert u_gen.get_index_by_str("txt2", ["txt1", "txt2"], True) == 1, (
        "Failed case: text str 'txt2', opt_list=['txt1', 'txt2'],"
        " minus one true.")


def test_clas_SafeDict_func__missing__1():
    assert "{key}".format_map(u_gen.SafeDict(dict())) == "{key}", (
        "Failed case: missing key should return the key inside braces '{key}'")
