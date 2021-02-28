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

def copy_sheet_in_new_workbook(workbook_path: str, sheet_to_copy_name: str, name_of_new_worbkook: str,name_of_new_sheet: str) -> None:
    workbook_to_copy = xw.Book(workbook_path)
    sheet_to_copy = workbook_to_copy.sheets[sheet_to_copy_name]
    copy_sheet_in_same_workbook(workbook_path=workbook_path,
                                sheet_to_copy_name=sheet_to_copy_name,
                                name_of_new_sheet=name_of_new_sheet)

    create_new_workbook(output_path=name_of_new_worbkook)
    new_worbkook = xw.Book(name_of_new_worbkook)

    new_sheet = new_worbkook.sheets[0]

    sheet_to_copy.api_copy_worksheet(before_=sheet_to_copy.api)

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
    workbook_path: str, sheet_to_modify_name: str, new_workbook_path: str, mapping:str
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

    # Todo: instead of creating new workbook, copy template workbook with formatting, to new_workbook_path
    create_new_workbook(output_path=new_workbook_path)


    formatted_signficant_dataframe = bring_pandas_dataframe_to_form_for_significant_mapping(dataframe=mapping_dataframe)

    write_pandas_dataframe_in_worksheet(dataframe=formatted_signficant_dataframe, workbook_path=new_workbook_path)


def bring_pandas_dataframe_to_form_for_significant_mapping(dataframe: pd.DataFrame) -> pd.DataFrame:

    columns = ['GL Acct', 'Name', 'PY 31.12.2019', 'PY Ref', 'CY 31.12.2020', 'CY Ref', 'Movement', 'Perc. Movement', 'Work Performed ref', 'Comments']
    singificant_pd = pd.DataFrame(columns=columns)

    singificant_pd['GL Acct']  = dataframe['GL Acct']
    singificant_pd['Name']  = dataframe['Name']
    singificant_pd['PY 31.12.2019'] = dataframe['PY 31.12.2019']
    singificant_pd['CY 31.12.2020'] = dataframe['CY 31.12.2020']

    singificant_pd['Movement'] = singificant_pd['CY 31.12.2020'] - singificant_pd['PY 31.12.2019']
    # todo: if denominator is 0, set Perc. Movement to 1.
    singificant_pd['Perc. Movement'] = singificant_pd['Movement'] / singificant_pd['PY 31.12.2019']

    return singificant_pd



def create_significant_leadsheets(
    workbook_path: str, sheet_to_modify_name: str, output_path: str, significant_mappings: List[str]
) -> None:
    for mapping in significant_mappings:
        create_leadsheet_given_mapping(
            workbook_path=workbook_path,
            sheet_to_modify_name=sheet_to_modify_name,
            new_workbook_path=f'{output_path}/{mapping}-leadsheet.xlsx',
            mapping=mapping,
        )


# def create_pandas_insignificant_dataframe()->pd.DataFrame -> pd.DataFrame:
#     """
#     This function creates the appropriate form for insignificant leadsheets
#     """


def create_insignificant_leadsheets(
    workbook_path: str, sheet_to_modify_name: str, output_path: str, insignificant_mappings: List[str]
) -> None:
    pd_df = create_pandas_dataframe_from_worksheet(
        workbook_path=workbook_path, sheet_to_modify_name=sheet_to_modify_name
    )

    create_new_workbook(output_path=output_path)

    insignificant_mappings_df = pd_df[pd_df['Mapping'].isin(insignificant_mappings)]

    write_pandas_dataframe_in_worksheet(dataframe=insignificant_mappings_df[COLUMNS_TO_USE], workbook_path=output_path)


def get_insignificant_mappings(dataframe: pd.DataFrame, significant_mappings: List[str]) -> List[str]:

    all_mappings = list(dataframe['Mapping'].unique())

    insignificant_mappings = []

    for element in all_mappings:
        if element not in significant_mappings:
            insignificant_mappings.append(element)

    return insignificant_mappings


# todo: add a column Significant mappings (list) or replace the input mapping
