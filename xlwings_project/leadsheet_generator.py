import xlwings as xw
import auditing_automation.excel_utils as excel_utils
import os

INPUT_WORKSHEET_NAME = 'Trial Balance'
SIGNIFICANT_MAPPINGS_COL = 'Significant Mappings'

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(THIS_DIR, 'data')

FORMATTED_TEMPLATE_WORKBOOK_PATH = os.path.join(THIS_DIR, 'leadsheet_format_template.xlsx')

LEADSHEETS_DIR = os.path.join(THIS_DIR, 'leadsheets')


def main():
    wb = xw.Book.caller()

    template_balance_sheet_path = str(wb.fullname)

    pd_df = excel_utils.create_pandas_dataframe_from_worksheet(
        workbook_path=template_balance_sheet_path, sheet_to_modify_name=INPUT_WORKSHEET_NAME
    )

    significant_mappings = excel_utils.get_significant_mappings(
        dataframe=pd_df, significant_mapping_col=SIGNIFICANT_MAPPINGS_COL
    )

    nonsignificant_mappings = excel_utils.get_nonsignificant_mappings(
        dataframe=pd_df, significant_mappings=significant_mappings
    )

    if not os.path.isdir(LEADSHEETS_DIR):
        os.mkdir(LEADSHEETS_DIR)

    excel_utils.create_significant_leadsheets(
        workbook_path=template_balance_sheet_path,
        sheet_to_modify_name=INPUT_WORKSHEET_NAME,
        output_path=LEADSHEETS_DIR,
        significant_mappings=significant_mappings,
        formatted_template_workbook=FORMATTED_TEMPLATE_WORKBOOK_PATH,
    )
    #
    excel_utils.create_nonsignificant_leadsheets(
        trial_balance_workbook_path=template_balance_sheet_path,
        trial_balance_sheet_name=INPUT_WORKSHEET_NAME,
        nonsignficant_workbook_path=f'{LEADSHEETS_DIR}/nonsignificant-leadsheets.xlsx',
        formatted_template_workbook=FORMATTED_TEMPLATE_WORKBOOK_PATH,
        nonsignificant_mappings=nonsignificant_mappings,
    )


if __name__ == "__main__":
    xw.Book("auditing_automation.xlam").set_mock_caller()
    main()
