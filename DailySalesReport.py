import pandas as pd
import datetime as dt
from xlsxwriter.utility import xl_rowcol_to_cell
import xlrd as xr
import win32com.client as win32
import mysql.connector as msql


#main function
def main():
    #run and save excel report
    run = dailyreport()

    #convert xlsx file to pdf
    file = convert_to_pdf()

    #open outlook and send email with pdf as attachment if the sum of the total sales is larger than 0
    if run > 0:
        sendEmail(file)
    else:
        pass

#function to run the daily sales report and save excel sheet
def dailyreport():

    #create db connection
    cnx = msql.connect(user='gone', password='fishing', port = '3305', host='192.168.88.88', database='fintronx_llc_live')

    #store sql file
    dailysales = open(r"C:\Users\dhumm\OneDrive\Documents\SQL_FB_Queries\dailySalesReport.sql")

    #run/read the sql file
    dsales = pd.read_sql_query(dailysales.read(), cnx)

    #find total amount of daily sales to return to main function
    total = dsales['Total_Price'].sum()

    #write to excel sheet and format columns and data accordingly, save and close book
    new_file = pd.ExcelWriter('C:\\Users\\dhumm\\automate\\Daily Sales.xlsx', datetime_format = 'mm/dd/yy', engine='xlsxwriter')
    dsales.to_excel(new_file, sheet_name = 'Daily Sales', index = False, startrow = 3)
    wb = new_file.book
    ws = new_file.sheets['Daily Sales']
    ws.set_zoom(90)
    center = wb.add_format({'align': 'center'})
    money_fmt = wb.add_format({'num_format': '$#,##0.00', 'align': 'center'})
    title_fmt = wb.add_format({'bold': True,'font_size':20, 'align': 'center', 'valign': 'vcenter'})
    total_fmt = wb.add_format({'bold': True, 'font_size': 14, 'align':'center', 'bg_color': 'yellow', 'top': 1})
    total_amt_fmt = wb.add_format({'bold': True, 'num_format': '$#,##0.00','font_size': 14, 'align':'center', 'bg_color': 'yellow', 'top': 1})
    ws.merge_range('A1:E3','Fintronx Daily Sales Report', title_fmt)
    ws.merge_range('F2:G2', 'Date: ' + dt.date.today().strftime("%m/%d/%Y"))
    ws.set_zoom(90)
    ws.set_column('A:C', 20, center)
    ws.set_column('D:E', 30, center)
    ws.set_column('F:F', 20, money_fmt)
    ws.set_column('G:G', 15, center)
    ws.write_formula('F'+str(len(dsales)+6), '=sum(F5:F'+str(len(dsales)+5)+')', total_amt_fmt)
    ws.write('E'+str(len(dsales)+6), 'Total For The Day', total_fmt)
    ws.set_landscape()
    ws.center_horizontally()
    ws.fit_to_pages(1,0)
    new_file.save()
    wb.close()

    return total

#converts xlsx file to pdf
def convert_to_pdf():
    #open excel wb from above, save as pdf, close excel
    xlapp = win32.DispatchEx('Excel.Application')
    xlwb = xlapp.Workbooks.Open('C:\\Users\\dhumm\\automate\\Daily Sales.xlsx')
    xlwb.ExportAsFixedFormat(0, "C:\\Users\\dhumm\\automate\\Daily Sales.pdf")
    xlwb.Close(True)

    attachment  = r"C:\Users\dhumm\automate\Daily Sales.pdf"

    return attachment

#function that sends automated email
def sendEmail(file):

    #send email to desired receipients
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0x0)
    #mail.To = 'dhummel@fintronxled.com'
    mail.To = '''droberson@fintronxled.com; dfaithful@fintronxled.com; neil.tolley@fintronxled.com;
    elucas@fintronxled.com;  swordsworth@cpfrm.com; pforbis@fintronxled.com; byates@cpfrm.com; clin@fintronxled.com;
    eadams@fintronxled.com; dhummel@fintronxled.com; dmartin@execdomain.com'''
    mail.Subject = 'Fintronx Daily Sales Report'
    mail.Body = 'See attached daily sales report.  Please let me know if you no longer wish to receive this email.'

    #attach excel report to the email

    mail.Attachments.Add(file)

    #send email
    mail.Send()


main()
