import pandas as pd
import datetime as dt
from xlsxwriter.utility import xl_rowcol_to_cell
import win32com.client as win32
import win32com.client.dynamic as win32d
import mysql.connector as msql


#function to run the weekly sales report
def weeklyreport():

    #create db connection
    cnx = msql.connect(user='gone', password='fishing', port = '3305', host='192.168.88.88', database='fintronx_llc_live')

    #store sql file
    weeklysales = open(r"C:\Users\dhumm\OneDrive\Documents\SQL_FB_Queries\weeklySalesReport.sql")

    #run/read the sql file
    wsales = pd.read_sql_query(weeklysales.read(), cnx)

    #write to excel sheet and format columns and data accordingly, save and close book
    new_file = pd.ExcelWriter('C:\\Users\\dhumm\\automate\\Weekly Sales For Week ' + dt.date.today().strftime("%W")+ '.xlsx', datetime_format = 'mm/dd/yy')
    wsales.to_excel(new_file, sheet_name = 'Weekly Sales For Week ' + dt.date.today().strftime("%W"), index = False, startrow = 3)
    wb = new_file.book
    ws = new_file.sheets['Weekly Sales For Week ' + dt.date.today().strftime("%W")]
    ws.set_zoom(90)
    center = wb.add_format({'align': 'center'})
    money_fmt = wb.add_format({'num_format': '$#,##0.00', 'align': 'center'})
    title_fmt = wb.add_format({'bold': True,'font_size':20, 'align': 'center', 'valign': 'vcenter'})
    ws.merge_range('A1:E3','Fintronx Weekly Sales Report', title_fmt)
    total_fmt = wb.add_format({'bold': True, 'font_size': 14, 'align':'center', 'bg_color': 'yellow', 'top': 1})
    total_amt_fmt = wb.add_format({'bold': True, 'num_format': '$#,##0.00','font_size': 14, 'align':'center', 'bg_color': 'yellow', 'top': 1})
    ws.merge_range('F2:G2', 'Date: ' + dt.date.today().strftime("%m/%d/%Y"))
    ws.set_zoom(90)
    ws.set_column('A:C', 20, center)
    ws.set_column('D:E', 30, center)
    ws.set_column('F:F', 20, money_fmt)
    ws.set_column('G:G', 15, center)
    ws.write_formula('F'+str(len(wsales)+6), '=sum(F5:F'+str(len(wsales)+5)+')', total_amt_fmt)
    ws.write('E'+str(len(wsales)+6), 'Total For The Week', total_fmt)
    ws.set_landscape()
    ws.center_horizontally()
    ws.fit_to_pages(1,0)
    new_file.save()
    new_file.close()

    #open excel wb from above, save as pdf, close excel
    xlapp = win32.DispatchEx('Excel.Application')

    xlwb = xlapp.Workbooks.Open('C:\\Users\\dhumm\\automate\\Weekly Sales For Week ' + dt.date.today().strftime("%W")+ '.xlsx')
    xlwb.ExportAsFixedFormat(0, "C:\\Users\\dhumm\\automate\\Weekly Sales For Week "+ dt.date.today().strftime("%W") + ".pdf")
    xlwb.Close(True)

    #send email to desired receipients
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    #mail.To = 'dhummel@fintronxled.com'
    mail.To = '''droberson@fintronxled.com; dfaithful@fintronxled.com; neil.tolley@fintronxled.com;
    elucas@fintronxled.com;  swordsworth@cpfrm.com; pforbis@fintronxled.com; byates@cpfrm.com; clin@fintronxled.com;
    eadams@fintronxled.com; dhummel@fintronxled.com; dmartin@execdomain.com'''
    mail.Subject = 'Fintronx Weekly Sales Report'
    mail.Body = 'See attached weekly sales report.  Please let me know if you no longer wish to receive this email.'

    #attach excel report to the email
    attachment  = r"C:\Users\dhumm\automate\Weekly Sales For Week "+ dt.date.today().strftime("%W") + ".pdf"
    mail.Attachments.Add(attachment)

    #send email
    mail.Send()


weeklyreport()
