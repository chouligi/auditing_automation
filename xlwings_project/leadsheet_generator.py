import xlwings as xw
import auditing_automation.excel_utils as excel_utils
import os

INPUT_WORKSHEET_NAME = 'Trial Balance'
INPUT_MAPPING_COL = 'Input - Mapping'
THIS_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(THIS_DIR, 'data')

TEST_SHEET_PATH = '/Users/Fosa/PythonProjects/git_tree/auditing_automation/xlwings_project/test_sheet.xlsx'

def main():
    excel_utils.create_leadsheet(workbook_path=TEST_SHEET_PATH,
                                 sheet_to_modify_name=INPUT_WORKSHEET_NAME,
                                 new_workbook_path=f'{THIS_DIR}/Cash-leadsheet.xlsx')


if __name__ == "__main__":
    xw.Book("auditing_automation.xlam").set_mock_caller()
    main()
