from insightly import Insightly
import pandas as pd
import datetime as dt
from xlsxwriter.utility import xl_rowcol_to_cell
import win32com.client as win32

#connect to insightly API
i = Insightly(apikey='6198d026-4015-40b6-a877-df96ecece0b1',version='2.2',debug=True)

#read all opportunities from insightly as a list of dictionaries
opps = i.read('opportunities', top = 500)

#read pipeline stages from instightly as list of dictionaries
stages = i.read('pipelinestages')

#read users from instightly as list of dictionaries
users = i.read('users')

#turn insightly data into a dataframe
opps_df = pd.DataFrame(opps)

#takes a dataframe and column, converts the column to a datetime type
def changeDateFormat(df, col):
    df[col] = pd.to_datetime(df[col])#.dt.strftime('%m/%d/%Y')
    return df[col]

#creates open report for pat and emails excel file
def open_orders_75_and_up(func, opps_df):

    #drop unneeded columns
    opps_df = opps_df.drop(['ACTUAL_CLOSE_DATE', 'DATE_CREATED_UTC', 'DATE_UPDATED_UTC', 'BID_AMOUNT','BID_CURRENCY','BID_DURATION','BID_TYPE','CAN_DELETE','CAN_EDIT','IMAGE_URL','LINKS',
    'OPPORTUNITY_DETAILS','OPPORTUNITY_STATE_REASON_ID', 'PIPELINE_ID','OWNER_USER_ID','TAGS','VISIBLE_TEAM_ID','VISIBLE_TO','VISIBLE_USER_IDS'], axis = 1)

    #list all columns in dataframe and find the ones that contain dates
    l_cols = list(opps_df)
    sub = 'DATE'
    date_cols = [s for s in l_cols if sub.lower() in s.lower()]

    #use changeDateFormat function to convert all columns that contain dates to proper format
    for col in date_cols:
        func(opps_df, col)


    #dataframe will contain only opportunities that have 75% and up chance at closing and whose state is set to OPEN
    opps_df = opps_df[opps_df['PROBABILITY']>=75.0]
    opps_df = opps_df[opps_df['OPPORTUNITY_STATE'] == 'OPEN']

    #fill in blank pipeline stages
    opps_df['STAGE_ID'].fillna('No Stage Noted', inplace = True)

    #replace the pipeline stage id from the number to the actual stage name
    for stage in stages:
        opps_df['STAGE_ID'].where(opps_df['STAGE_ID'] != stage['STAGE_ID'], stage['STAGE_NAME'], inplace = True)

    #replace the responsible user id from the number to the person first and last name
    for user in users:
        opps_df['RESPONSIBLE_USER_ID'].where(opps_df['RESPONSIBLE_USER_ID'] != user['USER_ID'], user['FIRST_NAME'] + ' '+ user['LAST_NAME'], inplace = True)

    #create empty city, state, product needs and qty needs lists to hold the data from custom fields from insightly
    #loop through the custom fields column in the dataframe to select a city, state, product needs and qty needs for each record, even if none exists
    #append each record's data to the lists
    city = [opp[0]['FIELD_VALUE'] if opp != [] else '' for opp in opps_df['CUSTOMFIELDS']]
    state = [opp[1]['FIELD_VALUE'] if opp != [] else '' for opp in opps_df['CUSTOMFIELDS']]
    prod_qty1 = [opp[2]['FIELD_VALUE'] if opp != [] else '' for opp in opps_df['CUSTOMFIELDS']]
    prod_wanted1 = [opp[3]['FIELD_VALUE'] if opp != [] else '' for opp in opps_df['CUSTOMFIELDS']]
    prod_qty2 = [opp[4]['FIELD_VALUE'] if opp != [] else '' for opp in opps_df['CUSTOMFIELDS']]
    prod_wanted2 = [opp[5]['FIELD_VALUE'] if opp != [] else '' for opp in opps_df['CUSTOMFIELDS']]
    recurring = ['Recurring' if opp == 5878778 else 'Non-Recurring'  for opp in opps_df['CATEGORY_ID']]

    #assign each list to new dataframe columns
    opps_df['CITY']=city
    opps_df['STATE']=state
    opps_df['QTY_1']=prod_qty1
    opps_df['PRODUCT_1']=prod_wanted1
    opps_df['QTY_2']=prod_qty2
    opps_df['PRODUCT_2']=prod_wanted2
    opps_df['CATEGORY_ID']=recurring

    #custom fields column no longer needed in dataframe once city and state are obtained and entered into new columns
    opps_df = opps_df.drop('CUSTOMFIELDS', axis=1)

    #sort dataframe by forecasted close date
    opps_df.sort_values(by=['FORECAST_CLOSE_DATE'], inplace = True)

    # BEGIN EXCEL WRITER FORMATTING - ANY COLUMN CHANGES ABOVE WILL NEED TO BE ADJUSTED HERE !!!!!!!

    #write to excel sheet and format columns and data accordingly
    new_file = pd.ExcelWriter('2018 Opportunities Forecasted.xlsx', datetime_format = 'mm/dd/yy')
    opps_df.to_excel(new_file, index=False, sheet_name='2018 Forecasted Opportunities')
    wb = new_file.book
    ws = new_file.sheets['2018 Forecasted Opportunities']
    ws.set_zoom(90)

    #define special formats
    mon_fmt = wb.add_format({'num_format': '$#,##0'})
    perc_fmt = wb.add_format({'num_format': '0"%"'})

    #set sizes and special formats
    ws.set_column('A:C', 20)
    ws.set_column('D:D',28)
    ws.set_column('E:E', 20)
    ws.set_column('F:F',20, mon_fmt)
    ws.set_column('G:G', 20, perc_fmt)
    ws.set_column('H:K', 20)
    ws.set_column('L:O', 20)

    #savefile
    new_file.save()

    #send email to desired receipients
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'dhummel@fintronxled.com'
    mail.Subject = 'Updated 2018 Insightly Opportunities'
    mail.Body = 'See attached Insightly Opportunity Report.  Please let me know if you no longer wish to receive this email.'

    #attach excel report to the email
    attachment  = r"C:\Users\dhumm\automate\2018 Opportunities Forecasted.xlsx"
    mail.Attachments.Add(attachment)

    #send email
    mail.Send()

open_orders_75_and_up(changeDateFormat, opps_df)
