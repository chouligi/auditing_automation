import openpyxl
from openpyxl.workbook.workbook import Workbook
from types import GeneratorType
import os
import xlwings as xw
import auditing_automation.python_utils as py_utils
import pandas as pd
from typing import List

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(THIS_DIR, 'data')

COLUMNS_TO_USE = ['GL Acct', 'Name', 'PY 31.12.2019', 'CY 31.12.2020', 'Mapping', 'Subcategory']


def load_xl_workbook(path: str) -> Workbook:
    return openpyxl.load_workbook(path)


def get_worksheet_values_from_workbook(wookrbook: Workbook, worksheet_name: str) -> GeneratorType:
    return wookrbook[worksheet_name].values


def copy_sheet_in_same_workbook(workbook_path: str, sheet_to_copy_name: str, name_of_new_sheet: str) -> None:
    workbook_to_copy = xw.Book(workbook_path)
    sheet_to_copy = workbook_to_copy.sheets[sheet_to_copy_name]

    # copy within the same sheet
    sheet_to_copy.api.copy_worksheet(after_=sheet_to_copy.api)

    copied_sheet = workbook_to_copy.sheets[1]

    copied_sheet.name = name_of_new_sheet


def copy_sheet_in_new_workbook(
    workbook_path: str, sheet_to_copy_name: str, name_of_new_workbook: str, name_of_new_sheet: str
) -> None:
    workbook_to_copy = xw.Book(workbook_path)
    sheet_to_copy = workbook_to_copy.sheets[sheet_to_copy_name]

    create_new_workbook(output_path=name_of_new_workbook)
    new_workbook = xw.Book(name_of_new_workbook)

    write_worksheet_in_new_workbook(
        workbook=new_workbook, sheet_to_copy=sheet_to_copy, name_of_new_sheet=name_of_new_sheet
    )


def copy_sheet_in_existing_workbook(
    workbook_path: str, sheet_to_copy_name: str, name_of_existing_workbook: str, name_of_new_sheet: str
) -> None:
    workbook_to_copy = xw.Book(workbook_path)
    sheet_to_copy = workbook_to_copy.sheets[sheet_to_copy_name]

    new_workbook = xw.Book(name_of_existing_workbook)

    write_worksheet_in_new_workbook(
        workbook=new_workbook, sheet_to_copy=sheet_to_copy, name_of_new_sheet=name_of_new_sheet
    )


def write_worksheet_in_new_workbook(workbook: xw.Book, sheet_to_copy: xw.Book.sheets, name_of_new_sheet: str):

    new_sheet = workbook.sheets[0]
    # Before_ if you want the sheet_to_copy to go before the new_sheet, after_ otherwise.
    sheet_to_copy.api.copy_worksheet(before_=new_sheet.api)
    new_sheet.name = name_of_new_sheet
    workbook.sheets[1].delete()


def create_new_workbook(output_path: str) -> None:
    new_workbook = xw.Book()
    new_workbook.save(output_path)


def create_pandas_dataframe_from_worksheet(workbook_path: str, sheet_to_modify_name: str) -> pd.DataFrame:
    workbook = load_xl_workbook(workbook_path)
    values = workbook[sheet_to_modify_name].values
    columns = py_utils.get_columns(values)

    return pd.DataFrame(values, columns=columns)


def create_pandas_leadsheet_given_mapping(dataframe: pd.DataFrame, mapping: str) -> pd.DataFrame:
    return dataframe[dataframe['Mapping'] == mapping]


def write_pandas_dataframe_in_worksheet(dataframe: pd.DataFrame, workbook_path: str):
    workbook = xw.Book(workbook_path)
    workbook.sheets[0].range('A1').options(index=False, header=True).value = dataframe


def create_leadsheet_given_mapping(
    workbook_path: str,
    sheet_to_modify_name: str,
    new_workbook_path: str,
    mapping: str,
    formatted_template_workbook: str,
) -> None:
    """
    This function reads a sheet from workbook_path, and creates a new workbook B with a subset
    of the sheet of workbook A.
    :return:
    """
    dataframe = create_pandas_dataframe_from_worksheet(
        workbook_path=workbook_path, sheet_to_modify_name=sheet_to_modify_name
    )

    mapping_dataframe = create_pandas_leadsheet_given_mapping(dataframe=dataframe, mapping=mapping)

    copy_sheet_in_new_workbook(
        workbook_path=formatted_template_workbook,
        sheet_to_copy_name='Leadsheet',
        name_of_new_workbook=new_workbook_path,
        name_of_new_sheet=convert_spaces_to_underscores(mapping),
    )

    formatted_signficant_dataframe = bring_pandas_dataframe_to_form_for_significant_mapping(dataframe=mapping_dataframe)

    write_pandas_dataframe_in_worksheet(dataframe=formatted_signficant_dataframe, workbook_path=new_workbook_path)

    # todo: find until which row there are characters and delete the format from the rest


def convert_spaces_to_underscores(string: str) -> str:
    return string.replace(' ', '_')


def bring_pandas_dataframe_to_form_for_significant_mapping(dataframe: pd.DataFrame) -> pd.DataFrame:

    columns = [
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
    singificant_pd = pd.DataFrame(columns=columns)

    singificant_pd['GL Acct'] = dataframe['GL Acct']
    singificant_pd['Name'] = dataframe['Name']
    singificant_pd['PY 31.12.2019'] = dataframe['PY 31.12.2019']
    singificant_pd['CY 31.12.2020'] = dataframe['CY 31.12.2020']

    singificant_pd['Movement'] = singificant_pd['CY 31.12.2020'] - singificant_pd['PY 31.12.2019']
    # todo: if denominator is 0, set Perc. Movement to 1.
    singificant_pd['Perc. Movement'] = singificant_pd['Movement'] / singificant_pd['PY 31.12.2019']

    return singificant_pd


def create_significant_leadsheets(
    workbook_path: str,
    sheet_to_modify_name: str,
    output_path: str,
    significant_mappings: List[str],
    formatted_template_workbook: str,
) -> None:
    for mapping in significant_mappings:
        create_leadsheet_given_mapping(
            workbook_path=workbook_path,
            sheet_to_modify_name=sheet_to_modify_name,
            new_workbook_path=f'{output_path}/{mapping}-leadsheet.xlsx',
            mapping=mapping,
            formatted_template_workbook=formatted_template_workbook,
        )


def create_nonsignificant_workbook(
    workbook_path: str, nonsignificant_mappings: List[str], formatted_template_workbook: str
) -> None:
    create_new_workbook(output_path=workbook_path)

    nonsignificant_workbook = xw.Book(workbook_path)

    copy_sheet_in_existing_workbook(
        workbook_path=formatted_template_workbook,
        sheet_to_copy_name='Leadsheet',
        name_of_existing_workbook=workbook_path,
        name_of_new_sheet=convert_spaces_to_underscores(nonsignificant_mappings[0]),
    )

    if len(nonsignificant_mappings) > 1:
        for mapping in nonsignificant_mappings[1:]:
            copy_sheet_in_same_workbook(
                workbook_path=workbook_path,
                sheet_to_copy_name=nonsignificant_workbook.sheets[0].name,
                name_of_new_sheet=convert_spaces_to_underscores(mapping),
            )


def create_nonsigificant_leadsheet_given_mapping(
    trial_balance_workbook_path: str, trial_balance_sheet: str, workbook_to_add_path: str, mapping: str
):
    """
    Using the data in the trial_balance_workbook and trial_balance_sheet, creates a new leadsheet in the
    workbook_to_add_path with name "mapping".
    """
    pd_df = create_pandas_dataframe_from_worksheet(
        workbook_path=trial_balance_workbook_path, sheet_to_modify_name=trial_balance_sheet
    )
    workbook_with_nonsignificant = xw.Book(workbook_to_add_path)
    mapping_dataframe = create_pandas_leadsheet_given_mapping(dataframe=pd_df, mapping=mapping)
    formatted_signficant_dataframe = bring_pandas_dataframe_to_form_for_significant_mapping(dataframe=mapping_dataframe)

    workbook_with_nonsignificant.sheets[convert_spaces_to_underscores(mapping)].range('A1').options(
        index=False, header=True
    ).value = formatted_signficant_dataframe


def create_nonsignificant_leadsheets(
    trial_balance_workbook_path: str,
    trial_balance_sheet_name: str,
    nonsignficant_workbook_path: str,
    formatted_template_workbook: str,
    nonsignificant_mappings: List[str],
) -> None:
    # trial_balance_pd = create_pandas_dataframe_from_worksheet(
    #    workbook_path=trial_balance_workbook_path, sheet_to_modify_name=trial_balance_sheet_name
    # )

    # nonsignificant_mappings = get_nonsignificant_mappings(dataframe=trial_balance_pd,
    #                            significant_mappings=significant_mappings)
    #
    create_nonsignificant_workbook(
        workbook_path=nonsignficant_workbook_path,
        nonsignificant_mappings=nonsignificant_mappings,
        formatted_template_workbook=formatted_template_workbook,
    )

    for mapping in nonsignificant_mappings:
        create_nonsigificant_leadsheet_given_mapping(
            trial_balance_workbook_path=trial_balance_workbook_path,
            trial_balance_sheet=trial_balance_sheet_name,
            workbook_to_add_path=nonsignficant_workbook_path,
            mapping=mapping,
        )


def get_nonsignificant_mappings(dataframe: pd.DataFrame, significant_mappings: List[str]) -> List[str]:

    all_mappings = list(dataframe['Mapping'].unique())

    nonsignificant_mappings = []

    for element in all_mappings:
        if element not in significant_mappings:
            nonsignificant_mappings.append(element)

    return nonsignificant_mappings


# todo: add a column Significant mappings (list) -> replace the input mapping
