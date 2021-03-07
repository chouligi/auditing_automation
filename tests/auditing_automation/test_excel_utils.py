from auditing_automation.excel_utils import (
    load_xl_workbook,
    get_worksheet_values_from_workbook,
    create_new_workbook,
    copy_sheet_in_same_workbook,
    create_leadsheet_given_mapping,
    create_pandas_dataframe_from_worksheet,
    write_pandas_dataframe_in_worksheet,
    create_pandas_leadsheet_given_mapping,
    create_significant_leadsheets,
    create_nonsignificant_leadsheets,
    get_nonsignificant_mappings,
    bring_pandas_dataframe_to_form_for_significant_mapping,
    copy_sheet_in_new_workbook,
    write_worksheet_in_new_workbook,
    convert_spaces_to_underscores,
    create_nonsignificant_workbook,
    create_nonsigificant_leadsheet_given_mapping,
)
import pandas as pd
from openpyxl.workbook.workbook import Workbook
import types
import xlwings as xw
import os
import shutil

SIGNIFICANT_DIR = 'signficant_leadsheets'
NONSIGNIFICANT_LEADSHEETS_PATH = 'nonsignificant-leadsheets.xlsx'


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


def test_copy_sheet_in_new_workbook_new_sheet_name_correct(test_formatted_leadsheet_template):
    sheet_to_copy_name = 'Leadsheet'
    name_of_new_sheet = 'Cash-Leadsheet'
    name_of_new_workbook = 'created-formatted-test-copy-leadsheet.xlsx'

    copy_sheet_in_new_workbook(
        workbook_path=test_formatted_leadsheet_template,
        sheet_to_copy_name=sheet_to_copy_name,
        name_of_new_workbook=name_of_new_workbook,
        name_of_new_sheet=name_of_new_sheet,
    )

    new_workbook = xw.Book(name_of_new_workbook)
    assert new_workbook.sheets[0].name == name_of_new_sheet

    os.remove(name_of_new_workbook)


def test_write_worksheet_in_new_workbook_name_is_correct(test_formatted_leadsheet_template):
    workbook_with_sheet_to_copy = xw.Book(test_formatted_leadsheet_template)
    new_workbook_name = 'created-formatted-test-write-leadsheet.xlsx'

    sheet_to_copy_name = 'Leadsheet'
    new_worksheet_name = 'Created Leadsheet'
    sheet_to_copy = workbook_with_sheet_to_copy.sheets[sheet_to_copy_name]
    create_new_workbook(new_workbook_name)

    workbook = xw.Book(new_workbook_name)
    write_worksheet_in_new_workbook(
        workbook=workbook, sheet_to_copy=sheet_to_copy, name_of_new_sheet=new_worksheet_name
    )

    assert workbook.sheets[0].name == new_worksheet_name
    os.remove(new_workbook_name)


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


def test_create_pandas_leadsheet_given_mapping_contains_cash_only(test_workbook):
    pd_df = create_pandas_dataframe_from_worksheet(workbook_path=test_workbook, sheet_to_modify_name='Trial Balance')
    cash_pd = create_pandas_leadsheet_given_mapping(dataframe=pd_df, mapping='Cash')

    assert cash_pd['Mapping'].unique() == ['Cash']


def test_create_pandas_leadsheet_given_mapping_contains_other_liabilities_only(test_workbook):
    pd_df = create_pandas_dataframe_from_worksheet(workbook_path=test_workbook, sheet_to_modify_name='Trial Balance')
    cash_pd = create_pandas_leadsheet_given_mapping(dataframe=pd_df, mapping='Other Liabilities')

    assert cash_pd['Mapping'].unique() == ['Other Liabilities']


def test_write_pandas_dataframe_in_worksheet(test_workbook):
    sheet_to_modify_name = 'Trial Balance'
    TEST_WORKBOOK_NAME = 'created_test_workbook.xlsx'

    formatted_dataframe = create_pandas_dataframe_from_worksheet(
        workbook_path=test_workbook, sheet_to_modify_name=sheet_to_modify_name
    )

    create_new_workbook(TEST_WORKBOOK_NAME)

    write_pandas_dataframe_in_worksheet(dataframe=formatted_dataframe, workbook_path=TEST_WORKBOOK_NAME)

    # todo: do assertion

    os.remove(TEST_WORKBOOK_NAME)


def test_create_leadsheet_given_mapping_cash(test_workbook, test_formatted_leadsheet_template):
    sheet_to_modify_name = 'Trial Balance'
    new_workbook_path = 'created-leadsheet-cash.xlsx'
    create_leadsheet_given_mapping(
        workbook_path=test_workbook,
        sheet_to_modify_name=sheet_to_modify_name,
        new_workbook_path=new_workbook_path,
        mapping='Cash',
        formatted_template_workbook=test_formatted_leadsheet_template,
    )

    assert os.path.exists(new_workbook_path)

    os.remove(new_workbook_path)


def test_create_leadsheet_given_mapping_name_with_spaces(test_workbook, test_formatted_leadsheet_template):
    sheet_to_modify_name = 'Trial Balance'
    new_workbook_path = 'created-leadsheet-name-with-spaces.xlsx'
    create_leadsheet_given_mapping(
        workbook_path=test_workbook,
        sheet_to_modify_name=sheet_to_modify_name,
        new_workbook_path=new_workbook_path,
        mapping='Worksheet Name With Spaces',
        formatted_template_workbook=test_formatted_leadsheet_template,
    )

    assert os.path.exists(new_workbook_path)

    os.remove(new_workbook_path)


def test_convert_spaces_to_underscores():

    assert convert_spaces_to_underscores('Trades And Other Receivables') == 'Trades_And_Other_Receivables'
    assert convert_spaces_to_underscores('Bla Bla Blablabla') == 'Bla_Bla_Blablabla'


def test_create_significant_leadsheets(test_workbook, test_formatted_leadsheet_template):

    sheet_to_modify_name = 'Trial Balance'

    pd_df = create_pandas_dataframe_from_worksheet(
        workbook_path=test_workbook, sheet_to_modify_name=sheet_to_modify_name
    )

    significant_mappings = pd_df['Mapping'].unique()

    assert not os.path.isdir(SIGNIFICANT_DIR)

    os.mkdir(SIGNIFICANT_DIR)

    create_significant_leadsheets(
        workbook_path=test_workbook,
        sheet_to_modify_name=sheet_to_modify_name,
        output_path=SIGNIFICANT_DIR,
        significant_mappings=significant_mappings,
        formatted_template_workbook=test_formatted_leadsheet_template,
    )

    assert os.path.isdir(SIGNIFICANT_DIR)

    assert os.path.exists(os.path.join(SIGNIFICANT_DIR, 'Cash-leadsheet.xlsx'))

    shutil.rmtree(SIGNIFICANT_DIR)


def test_create_nonsigificant_leadsheet_given_mapping(test_workbook, test_formatted_leadsheet_template):
    sheet_to_modify_name = 'Trial Balance'

    nonsignificant_mappings = ['Other Liabilities', 'Trade And Other Receivables']

    mapping = 'Trade And Other Receivables'

    workbook_to_add_path = 'test-nonsignificant-leadsheets-given-mapping.xlsx'

    create_nonsignificant_workbook(
        workbook_path=workbook_to_add_path,
        nonsignificant_mappings=nonsignificant_mappings,
        formatted_template_workbook=test_formatted_leadsheet_template,
    )

    create_nonsigificant_leadsheet_given_mapping(
        trial_balance_workbook_path=test_workbook,
        trial_balance_sheet=sheet_to_modify_name,
        workbook_to_add_path=workbook_to_add_path,
        mapping=mapping,
    )

    nonsignificant_workbook = xw.Book(workbook_to_add_path)
    nonsignificant_workbook.save(path=workbook_to_add_path)
    # todo: use this way to load pandas dataframes!
    excel_file = pd.ExcelFile(workbook_to_add_path)
    pd_df = excel_file.parse(convert_spaces_to_underscores(mapping))

    assert len(pd_df) > 2
    os.remove(workbook_to_add_path)


def test_create_nonsignificant_leadsheets(test_workbook, test_formatted_leadsheet_template):
    sheet_to_modify_name = 'Trial Balance'
    nonsignificant_workbook_path = 'test-nonsignificant-leadsheets.xlsx'

    nonsignificant_mappings = ['Other Liabilities', 'Trade And Other Receivables']

    create_nonsignificant_leadsheets(
        trial_balance_workbook_path=test_workbook,
        trial_balance_sheet_name=sheet_to_modify_name,
        nonsignficant_workbook_path=nonsignificant_workbook_path,
        formatted_template_workbook=test_formatted_leadsheet_template,
        nonsignificant_mappings=nonsignificant_mappings,
    )

    assert os.path.isfile(nonsignificant_workbook_path)
    os.remove(nonsignificant_workbook_path)


def test_create_nonsignificant_workbook_path_and_number_of_sheets(test_workbook, test_formatted_leadsheet_template):
    workbook_path = 'test-path-nonsignificant-leadsheets.xlsx'

    nonsignificant_mappings = ['Other Liabilities', 'Trade And Other Receivables']

    create_nonsignificant_workbook(
        workbook_path=workbook_path,
        nonsignificant_mappings=nonsignificant_mappings,
        formatted_template_workbook=test_formatted_leadsheet_template,
    )

    workbook = xw.Book(workbook_path)
    assert os.path.exists(workbook_path)
    sheet_names = [sheet.name for sheet in workbook.sheets]
    assert len(sheet_names) == len(nonsignificant_mappings)

    os.remove(workbook_path)


def test_get_nonsignificant_mappings(test_workbook):
    sheet_to_modify_name = 'Trial Balance'
    significant_mappings = ['Cash']

    pd_df = create_pandas_dataframe_from_worksheet(
        workbook_path=test_workbook, sheet_to_modify_name=sheet_to_modify_name
    )

    nonsignificant_mappings = get_nonsignificant_mappings(pd_df, significant_mappings)

    assert nonsignificant_mappings == ['Trade And Other Receivables', 'Other Liabilities']


def test_get_nonsignificant_mappings_none_remaining(test_workbook):
    sheet_to_modify_name = 'Trial Balance'
    significant_mappings = ['Cash', 'Trade And Other Receivables', 'Other Liabilities']

    pd_df = create_pandas_dataframe_from_worksheet(
        workbook_path=test_workbook, sheet_to_modify_name=sheet_to_modify_name
    )

    nonsignificant_mappings = get_nonsignificant_mappings(pd_df, significant_mappings)

    assert nonsignificant_mappings == []


def test_bring_pandas_dataframe_to_form_for_significant_mapping_is_pandas_df(test_workbook):
    sheet_to_modify_name = 'Trial Balance'

    pd_df = create_pandas_dataframe_from_worksheet(
        workbook_path=test_workbook, sheet_to_modify_name=sheet_to_modify_name
    )

    significant_form_pd = bring_pandas_dataframe_to_form_for_significant_mapping(dataframe=pd_df)

    assert isinstance(significant_form_pd, pd.DataFrame)


def test_bring_pandas_dataframe_to_form_for_significant_mapping_contains_proper_columns(test_workbook):
    sheet_to_modify_name = 'Trial Balance'

    pd_df = create_pandas_dataframe_from_worksheet(
        workbook_path=test_workbook, sheet_to_modify_name=sheet_to_modify_name
    )
    expected_columns = [
        'GL Acct',
        'Name',
        'PY 31.12.2019',
        'PY Ref',
        'CY 31.12.2020',
        'CY Ref',
        'Movement',
        'Perc. Movement',
        'Work Performed ref',
        'Comments',
    ]
    significant_form_pd = bring_pandas_dataframe_to_form_for_significant_mapping(dataframe=pd_df)
    columns = [col for col in significant_form_pd.columns]

    assert columns == expected_columns


def test_bring_pandas_dataframe_to_form_for_significant_mapping_movement_computed_correctly(test_workbook):
    return


def test_bring_pandas_dataframe_to_form_for_significant_mapping_perc_movement_computed_correctly(test_workbook):
    return


def test_bring_pandas_dataframe_to_form_for_significant_mapping_perc_movement_returns_1_when_denom_is_0(test_workbook):
    return
