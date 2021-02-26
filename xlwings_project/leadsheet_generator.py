import xlwings as xw
import auditing_automation.excel_utils as excel_utils
import os

INPUT_WORKSHEET_NAME = 'Trial Balance'
INPUT_MAPPING_COL = 'Input - Mapping'
THIS_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(THIS_DIR, 'data')

TEST_SHEET_PATH = os.path.join(THIS_DIR, 'test_sheet.xlsx')


SIGNIFICANT_MAPPINGS = ['Cash', 'Other Liabilities']
INSIGNIFICANT_MAPPINGS = ['Trade And Other Receivables']


def main():
    excel_utils.create_significant_leadsheets(
        workbook_path=TEST_SHEET_PATH,
        sheet_to_modify_name=INPUT_WORKSHEET_NAME,
        output_path=THIS_DIR,
        significant_mappings=SIGNIFICANT_MAPPINGS,
    )

    excel_utils.create_insignificant_leadsheets(
        workbook_path=TEST_SHEET_PATH,
        sheet_to_modify_name=INPUT_WORKSHEET_NAME,
        output_path=f'{THIS_DIR}/insignificant-leadsheets.xlsx',
        insignificant_mappings=INSIGNIFICANT_MAPPINGS,
    )


if __name__ == "__main__":
    xw.Book("auditing_automation.xlam").set_mock_caller()
    main()
