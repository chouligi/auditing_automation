import xlwings as xw
import pandas as pd
import openpyxl

INPUT_WORKSHEET_NAME = 'Trial Balance'
INPUT_MAPPING_COL = 'Input - Mapping'


def load_xl_workbook(path: str):
    return openpyxl.load_workbook(path)


def get_columns(data):
    return next(data)[0:]


def main():
    wb = xw.Book.caller()

    trial_balance_input_sheet = wb.sheets[0]

    df = load_xl_workbook(
        path='/Users/Fosa/PythonProjects/git_tree/auditing_automation/xlwings_project/test_sheet.xlsx')

    sheet = df
    values = df[INPUT_WORKSHEET_NAME].values
    # columns = next(data)[0:]

    columns = get_columns(values)
    p_df = pd.DataFrame(values, columns=columns)

    path = '/Users/Fosa/PythonProjects/git_tree/auditing_automation/xlwings_project/'

    required_mapping_type = p_df[INPUT_MAPPING_COL][0]
    formatted_dataframe = p_df[p_df['Mapping'] == required_mapping_type]

    # sheet = wb.sheets['Trial Balance']

    new_workbook = xw.Book()
    # add the formated dataframe in first sheet
    new_workbook.sheets[0].range('A1').options(index=False, header=True).value = formatted_dataframe
    # store the workbook as

    new_workbook.save(path=f"{path}{required_mapping_type}-leadsheet.xlsx")


if __name__ == "__main__":
    xw.Book("leadsheet_generator.xlam").set_mock_caller()
    main()