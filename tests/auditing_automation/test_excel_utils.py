from auditing_automation.excel_utils import (
    load_xl_workbook,
    get_worksheet_values_from_workbook,
    create_new_workbook,
    copy_sheet_in_same_workbook,
    create_leadsheet,
    create_pandas_dataframe_from_worksheet,
    write_dataframe_in_worksheet,
)
import pandas as pd
from openpyxl.workbook.workbook import Workbook
import types
import xlwings as xw
import os


def test_load_xl_workbook_is_openpyxl_workbookw(test_workbook):
    imported_workbook = load_xl_workbook(path=test_workbook)
    assert type(imported_workbook) == Workbook


def test_get_worksheet_values_from_workbook(test_workbook):
    imported_workbook = load_xl_workbook(path=test_workbook)

    worksheet_values = get_worksheet_values_from_workbook(imported_workbook, 'Trial Balance')

    assert isinstance(worksheet_values, types.GeneratorType)


def test_get_worksheet_values_from_workbook_contain_data(test_workbook):
    imported_workbook = load_xl_workbook(path=test_workbook)

    worksheet_values = get_worksheet_values_from_workbook(imported_workbook, 'Trial Balance')
    pd_df = pd.DataFrame(worksheet_values)

    assert len(pd_df) > 0


def test_copy_sheet_in_same_workbook(test_workbook):
    sheet_to_copy_name = 'Trial Balance'
    name_of_new_sheet = 'Copied Trial Balance'

    copy_sheet_in_same_workbook(
        workbook_path=test_workbook, sheet_to_copy_name=sheet_to_copy_name, name_of_new_sheet=name_of_new_sheet
    )

    workbook = xw.Book(test_workbook)

    sheet_names = [sheet.name for sheet in workbook.sheets]
    assert name_of_new_sheet in sheet_names

    for sheet in workbook.sheets:
        if sheet.name != sheet_to_copy_name:
            sheet.delete()


def test_create_new_workbook():
    test_name = 'test-leadsheet.xlsx'
    create_new_workbook(output_path=test_name)
    os.remove(test_name)


def test_create_pandas_dataframe_from_worksheet_is_pandas_df(test_workbook):

    pd_df = create_pandas_dataframe_from_worksheet(workbook_path=test_workbook, sheet_to_modify_name='Trial Balance')
    assert isinstance(pd_df, pd.DataFrame)


def test_create_pandas_dataframe_from_worksheet_is_not_empty(test_workbook):
    pd_df = create_pandas_dataframe_from_worksheet(workbook_path=test_workbook, sheet_to_modify_name='Trial Balance')
    assert len(pd_df) > 0


def test_write_dataframe_in_worksheet(test_workbook):
    sheet_to_modify_name = 'Trial Balance'
    TEST_WORKBOOK_NAME = 'created_test_workbook.xlsx'

    formatted_dataframe = create_pandas_dataframe_from_worksheet(
        workbook_path=test_workbook, sheet_to_modify_name=sheet_to_modify_name
    )

    create_new_workbook(TEST_WORKBOOK_NAME)

    write_dataframe_in_worksheet(dataframe=formatted_dataframe, workbook_path=TEST_WORKBOOK_NAME)

    os.remove(TEST_WORKBOOK_NAME)

def test_create_leadsheet(test_workbook):
    sheet_to_modify_name = 'Trial Balance'
    new_workbook_path = 'created-leadsheet.xlsx'
    create_leadsheet(
        workbook_path=test_workbook, sheet_to_modify_name=sheet_to_modify_name, new_workbook_path=new_workbook_path
    )
    os.remove(new_workbook_path)

