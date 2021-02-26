from auditing_automation.excel_utils import (
    load_xl_workbook,
    get_worksheet_values_from_workbook,
    create_new_workbook,
    copy_sheet_in_same_workbook,
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
    test_name = 'test'
    create_new_workbook(output_path=test_name)
    os.remove(f'{test_name}-leadsheet.xlsx')
