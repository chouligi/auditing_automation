import pandas as pd
import openpyxl
balance_sheet = 'Test_auditing_notebook.xlsx'

x = pd.read_excel(balance_sheet, index_col=None, header=None, engine='openpyxl')