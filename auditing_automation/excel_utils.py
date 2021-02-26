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

COLUMNS_OF_INTEREST = ['GL Acct', 'Name', 'PY 31.12.2019', 'CY 31.12.2020', 'Mapping', 'Subcategory']


def load_xl_workbook(path: str) -> Workbook:
    return openpyxl.load_workbook(path)


def get_worksheet_values_from_workbook(wookrbook: Workbook, worksheet_name: str) -> GeneratorType:
    return wookrbook[worksheet_name].values


def copy_sheet_in_same_workbook(workbook_path: str, sheet_to_copy_name: str, name_of_new_sheet: str):
    workbook_to_copy = xw.Book(workbook_path)
    sheet_to_copy = workbook_to_copy.sheets[sheet_to_copy_name]

    # copy within the same sheet
    sheet_to_copy.api.copy_worksheet(after_=sheet_to_copy.api)

    copied_sheet = workbook_to_copy.sheets[1]

    copied_sheet.name = name_of_new_sheet


def create_new_workbook(output_path: str):
    new_workbook = xw.Book()
    new_workbook.save(output_path)


# def create_pandas_dataframe_from_worksheet(workbook_path: str, sheet_to_modify_name: str):
#    INPUT_MAPPING_COL = 'Input - Mapping'
#
#    workbook = load_xl_workbook(workbook_path)
#    values = workbook[sheet_to_modify_name].values
#    columns = py_utils.get_columns(values)
#    p_df = pd.DataFrame(values, columns=columns)
#    required_mapping_type = p_df[INPUT_MAPPING_COL][0]
#
#    return p_df[p_df['Mapping'] == required_mapping_type]


def create_pandas_dataframe_from_worksheet(workbook_path: str, sheet_to_modify_name: str):
    workbook = load_xl_workbook(workbook_path)
    values = workbook[sheet_to_modify_name].values
    columns = py_utils.get_columns(values)

    return pd.DataFrame(values, columns=columns)


def create_pandas_leadsheet_given_mapping(dataframe: pd.DataFrame, mapping: str):
    return dataframe[dataframe['Mapping'] == mapping]


def write_pandas_dataframe_in_worksheet(dataframe: pd.DataFrame, workbook_path: str):
    workbook = xw.Book(workbook_path)
    workbook.sheets[0].range('A1').options(index=False, header=True).value = dataframe


def create_leadsheet_given_mapping(
    workbook_path: str, sheet_to_modify_name: str, new_workbook_path: str, mapping='str'
):
    """
    This function reads a sheet from workbook_path, and creates a new workbook B with a subset
    of the sheet of workbook A.
    :return:
    """
    formatted_dataframe = create_pandas_dataframe_from_worksheet(
        workbook_path=workbook_path, sheet_to_modify_name=sheet_to_modify_name
    )

    mapping_dataframe = create_pandas_leadsheet_given_mapping(dataframe=formatted_dataframe, mapping=mapping)

    create_new_workbook(output_path=new_workbook_path)
    write_pandas_dataframe_in_worksheet(dataframe=mapping_dataframe, workbook_path=new_workbook_path)


def create_significant_leadsheets(
    workbook_path: str, sheet_to_modify_name: str, output_path: str, significant_mappings: List
):
    for mapping in significant_mappings:
        create_leadsheet_given_mapping(
            workbook_path=workbook_path,
            sheet_to_modify_name=sheet_to_modify_name,
            new_workbook_path=f'{output_path}/{mapping}-leadsheet.xlsx',
            mapping=mapping,
        )

    # def create_pandas_insignificant_dataframe()->pd.DataFrame:
    """
    This function creates the appropriate form for insignificant leadsheets
    """


def create_insignificant_leadsheets(
    workbook_path: str, sheet_to_modify_name: str, output_path: str, insignificant_mappings: List
):
    pd_df = create_pandas_dataframe_from_worksheet(
        workbook_path=workbook_path, sheet_to_modify_name=sheet_to_modify_name
    )

    create_new_workbook(output_path=output_path)

    insignificant_mappings_df = pd_df[pd_df['Mapping'].isin(insignificant_mappings)]

    write_pandas_dataframe_in_worksheet(dataframe=insignificant_mappings_df, workbook_path=output_path)
