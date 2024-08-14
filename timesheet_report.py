import pandas as pd
from openpyxl.comments import Comment
from openpyxl import load_workbook

time_sheet_df = pd.read_excel('./Timelogs.xlsx',sheet_name='Sheet 1')
time_sheet_df['Employee Name'] = time_sheet_df['First name'] + ' ' + time_sheet_df['Last name']

## Change as per work item you want
time_sheet_df = time_sheet_df[time_sheet_df['Work Item'] == 'Not Applicable']

time_sheet_df[['Hrs','Min']] = time_sheet_df['Hours(HH:MM)'].str.split(':', expand=True)
time_sheet_df['Min'] = time_sheet_df['Min'].astype(int)/60
time_sheet_df['Hours'] = time_sheet_df['Min'] + time_sheet_df['Hrs'].astype(float)
time_sheet_df = time_sheet_df.loc[:, ~time_sheet_df.columns.isin([
    'Employee id', 'Client Name','Email ID',
    'Last name','From time','First name',
    'To time','Timer Intervals','Hour(s)',
    'Billing Status','Approval Status','Description','Work Item',
    'Hours(HH:MM)','Hrs','Min'
    ])]
time_sheet_df.rename(columns={'Date' : ' '},inplace=True)
time_sheet_df = time_sheet_df.pivot_table(index=' ', columns=['Project Name','Job Name','Employee Name'], values='Hours',aggfunc='sum')
time_sheet_df.fillna(0,inplace=True)
time_sheet_df.to_excel('./New Reports/Report_Project_Task.xlsx')