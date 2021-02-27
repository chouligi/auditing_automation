import xlwings as xw
import auditing_automation.excel_utils as excel_utils
import os

INPUT_WORKSHEET_NAME = 'Trial Balance'
INPUT_MAPPING_COL = 'Input - Mapping'
THIS_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(THIS_DIR, 'data')

SIGNIFICANT_MAPPINGS = ['Cash', 'Other Liabilities', 'Trade And Other Receivables']


def main():
    wb = xw.Book.caller()

    template_balance_sheet_path = str(wb.fullname)

    pd_df = excel_utils.create_pandas_dataframe_from_worksheet(
        workbook_path=template_balance_sheet_path, sheet_to_modify_name=INPUT_WORKSHEET_NAME
    )

    insignificant_mappings = excel_utils.get_insignificant_mappings(
        dataframe=pd_df, significant_mappings=SIGNIFICANT_MAPPINGS
    )

    excel_utils.create_significant_leadsheets(
        workbook_path=template_balance_sheet_path,
        sheet_to_modify_name=INPUT_WORKSHEET_NAME,
        output_path=THIS_DIR,
        significant_mappings=SIGNIFICANT_MAPPINGS,
    )

    excel_utils.create_insignificant_leadsheets(
        workbook_path=template_balance_sheet_path,
        sheet_to_modify_name=INPUT_WORKSHEET_NAME,
        output_path=f'{THIS_DIR}/insignificant-leadsheets.xlsx',
        insignificant_mappings=insignificant_mappings,
    )


if __name__ == "__main__":
    xw.Book("auditing_automation.xlam").set_mock_caller()
    main()
