import pandas as pd
from openpyxl.comments import Comment
from openpyxl import load_workbook

time_sheet_df = pd.read_excel('./final.xlsx',sheet_name='Sheet1')
time_sheet_df = time_sheet_df.transpose()
time_sheet_df.reset_index(drop=True,inplace=True)
time_sheet_df.drop(columns=[0,2],inplace=True)
time_sheet_df.columns = time_sheet_df.iloc[0]
time_sheet_df = time_sheet_df[1:]
final_df = time_sheet_df.groupby(['Employee Name']).sum().transpose()
final_df.index.name = 'Date'
final_df.to_excel('Report_WO_Project_Task.xlsx')