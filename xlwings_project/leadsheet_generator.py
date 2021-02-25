import xlwings as xw
import auditing_automation.utils as utils
import os

INPUT_WORKSHEET_NAME = 'Trial Balance'
INPUT_MAPPING_COL = 'Input - Mapping'
THIS_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(THIS_DIR, 'data')


def main():
    utils.copy_sheet_in_same_workbook(
        workbook_path='/Users/Fosa/PythonProjects/git_tree/auditing_automation/xlwings_project/test_sheet.xlsx',
        sheet_to_copy_name=INPUT_WORKSHEET_NAME,
        name_of_new_sheet='My New Sheet')


if __name__ == "__main__":
    xw.Book("auditing_automation.xlam").set_mock_caller()
    main()
