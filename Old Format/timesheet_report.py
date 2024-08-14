import pandas as pd
from openpyxl.comments import Comment
from openpyxl import load_workbook

time_sheet_df = pd.read_excel('tp.xlsx',sheet_name='sheet2')
time_sheet_df.drop(columns=['Total'],inplace=True)
time_sheet_df = time_sheet_df.loc[:, ~time_sheet_df.columns.isin(['Employee Id', 'Client Name','Email ID'])]
time_sheet_df = time_sheet_df.drop(time_sheet_df.index[len(time_sheet_df) - 1])
time_sheet_df = time_sheet_df.groupby(['Project Name','Job Name','Employee Name']).sum()
timesheet_transposed_df = time_sheet_df.transpose()
timesheet_transposed_df.to_excel('/Reports/Report_Project_Task.xlsx')