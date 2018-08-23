import pyodbc
import pandas as pd
import datetime
import os
import numpy as np
import win32com.client as win32


SalesReps = [('MBM', 'Myrick', 'myrick@stingerchemicals.com')
        , ('JST', 'Takkie', 'takkie@stingerchemicals.com; andrea@stingerchemicals.com; jon@stingerchemicals.com')
        , ('FL','Frank','frank@stingerchemicals.com')
        , ('THF','Tim Floyd','tfloyd@stingerchemicals.com;leigh@stingerchemicals.com')
        ,('GR','Greg','greg@stingerchemicals.com')
        , ('AV','Alton','alton@stingerchemicals.com'),('AR','Albert','albertr@stingerchemicals.com')
		, ('LB','Larry Bale','larryb@stingerchemicals.com; noah@stingerchemicals.com')]

CCEmails = 'warren@stingerchemicals.com; stu@stingerchemicals.com; fritz@stingerchemicals.com'

cn = pyodbc.connect('DSN=QuickBooks Data;')

#store procedure in QB to query Open Invoices report as of today
sql = """sp_report OpenInvoices show TxnType, Name, Date, RefNumber, PONumber, Terms, DueDate, Aging, SalesRep, OpenBalance
            parameters DateMacro = 'Today'"""

#read data from sql and connection
data = pd.read_sql(sql,cn)

#convert data from above to Pandas Dataframe
df = pd.DataFrame(data)


#loop through list of sales rep
for RepInitial, RepFullName, RepEmail in SalesReps:

    #filter the dataframe to pull data of each sales rep
    df2 = df[df.SalesRep == RepInitial]

    #df2 = pd.pivot_table(df2, index = ['TxnType', 'Name', 'Date', 'RefNumber', 'PONumber', 'Terms', 'DueDate', 'Aging', 'SalesRep'], values = ['OpenBalance'], aggfunc = [np.sum], fill_value=0)

    #filename is the output Excel file of each sales rep
    var_output_Excel_file = RepFullName + ' Open Invoices ' + str(datetime.date.today()) + '.xlsx'

    #filepath is the path to the filename above,
    #os.getcwd() is to get current dictory of Python script
    var_output_Excel_path = os.getcwd() + '\\' + var_output_Excel_file


    #initiate a write object using file name
    writer = pd.ExcelWriter(var_output_Excel_file, engine='xlsxwriter')

    #write the dataframe df2 to excel
    df2.to_excel(writer, sheet_name= RepFullName, startcol=0, startrow=0, index=False, header=True)

    #using pivot table to subtotal open balance of each customers, create a new worksheet named 'Summary'
    df3 = pd.pivot_table(df2, index=['Name'], values=['OpenBalance'], aggfunc=[np.sum], fill_value=0)
    df3.to_excel(writer, sheet_name='Summary', startcol=0, startrow=0, index=True, header=True)

    ## this function is to format the column Open balance
    workbook = writer.book
    worksheet1 = writer.sheets[RepFullName]

    customer_name_width = 18

    for row in df2['Name']:
        if customer_name_width < len(row):
            customer_name_width = len(row)  #find the longest Customer Name, to set the width of column Name later


    format = workbook.add_format()
    format.set_num_format('#,##0.00') ## format cell as number with commas
    format.set_bold() ## format bold for the cell


    worksheet1.set_column('J:J', 18, format) # apply the format to column J, set column width = 18
    worksheet1.set_column('A:J',customer_name_width) # set column A -> J width = longest customer's name
    worksheet1.freeze_panes(1, 0)  #freeze the top row

    worksheet2 = writer.sheets['Summary']
    worksheet2.set_column('B:B',18,format)  #apply number format to column
    worksheet2.set_column('A:A',30)         #apply number format to column



    #save output Excel file
    writer.save()



#this function is to email the output Excel file to each sales rep
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = RepEmail  #change this line to change receipient's emails
    mail.CC = CCEmails
    mail.Subject = RepFullName + ' Open Invoice as of ' + str(datetime.date.today())
    mail.Body = 'Message body'
    mail.HTMLBody = '<h2>This is Unpaid Invoices of ' + RepFullName + ' customers</h2>'

    mail.Attachments.Add(var_output_Excel_path)
    mail.Send()
    os.remove(var_output_Excel_path)       #delete the excel file after email sent

#close ODBC connection.
cn.close()






