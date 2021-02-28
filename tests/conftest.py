import os
import pytest

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
DATASETS_DIR = os.path.join(THIS_DIR, 'sample_data')


@pytest.fixture(scope='module')
def test_workbook() -> str:
    return os.path.join(DATASETS_DIR, 'test_workbook.xlsx')


@pytest.fixture(scope='module')
def test_formatted_leadsheet_template() -> str:
    return os.path.join(DATASETS_DIR, 'test_formatted_leadsheet_template.xlsx')
