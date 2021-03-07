from auditing_automation.excel_utils import load_xl_workbook, get_worksheet_values_from_workbook
from auditing_automation.python_utils import get_columns


def test_get_columns(test_workbook):
    imported_workbook = load_xl_workbook(path=test_workbook)
    worksheet_values = get_worksheet_values_from_workbook(imported_workbook, 'Trial Balance')

    columns = get_columns(worksheet_values)
    expected_output = (
        'GL Acct',
        'Name',
        'PY 31.12.2019',
        'CY 31.12.2020',
        'Mapping',
        'Subcategory',
        None,
        'Significant Mappings',
    )
    assert columns == expected_output
