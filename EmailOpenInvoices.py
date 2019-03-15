import pyodbc
import pandas as pd
import datetime
import os
import numpy as np
import win32com.client as win32



SalesReps = [('MBM', 'Myrick', 'myrick@stingerchemicals.com')
        , ('JST', 'Takkie', 'takkie@stingerchemicals.com; jon@stingerchemicals.com')
        , ('FL','Frank','frank@stingerchemicals.com')
        , ('THF','Tim Floyd','tfloyd@stingerchemicals.com')
        ,('GR','Greg','greg@stingerchemicals.com')
        , ('AV','Alton','alton@stingerchemicals.com'),('AR','Albert','albertr@stingerchemicals.com')
		, ('LB','Larry Bale','larryb@stingerchemicals.com; noah@stingerchemicals.com')
        , ('JD','Joey','joeyd@stingerchemicals.com')]


# SalesReps = [('FL','Frank','frank@stingerchemicals.com')]


CCEmails = 'warren@stingerchemicals.com; stu@stingerchemicals.com; fritz@stingerchemicals.com; \
leigh@stingerchemicals.com; bridget@stingerchemicals.com; \
andrea@stingerchemicals.com; kimberly@stingerchemicals.com; yvonne@stingerchemicals.com'

cn = pyodbc.connect('DSN=QuickBooks Data;')

#store procedure in QB to query Open Invoices report as of today
sql = """sp_report OpenInvoices show TxnType, Name, Date, RefNumber, PONumber, Terms, DueDate, Aging, SalesRep, OpenBalance
            parameters DateMacro = 'Today'"""

#read data from sql and connection
data = pd.read_sql(sql,cn)

#convert data from above to Pandas Dataframe
df = pd.DataFrame(data)
 
df['Aging'] = df['Aging'].fillna('')
#df['OpenBalance'] = df['OpenBalance'].map("{0:,.2f}".format)



#loop through list of sales rep
for RepInitial, RepFullName, RepEmail in SalesReps:

    #filter the dataframe to pull data of each sales rep
    df2 = df[df.SalesRep == RepInitial]

    #filename is the output Excel file of each sales rep
    var_output_Excel_file = RepFullName + ' Open Invoices ' + str(datetime.date.today()) + '.xlsx'

    #filepath is the path to the filename above,
    #BASE_DIR is to get current dictory of Python script
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    var_output_Excel_path = os.path.join(BASE_DIR,var_output_Excel_file)
    
  

    #initiate a write object using file name
    writer = pd.ExcelWriter(var_output_Excel_path, engine='xlsxwriter')

    #write the dataframe df2 to excel
    df2.to_excel(writer, sheet_name= RepFullName, startcol=0, startrow=0, index=False, header=True)    
    
    #this function is to hightlight rows base on the aging column.
    def highlight_pastdueinvoice(s):        
        if s.Aging != '' and s.Aging >= 30 and s.Aging < 60:
            return ['background-color: #007bff; font-weight: bold; color: white'] * s.size
        elif s.Aging != '' and s.Aging >= 60 and s.Aging < 90:
            return ['background-color: #ffc107; font-weight: bold'] * s.size    
        elif s.Aging != '' and s.Aging >= 90:
            return ['background-color: #dc3545; font-weight: bold; color: white'] * s.size  
        else:
            return ['background-color: white'] * s.size
    
    html_string = (df2.style.format({'OpenBalance':"{0:,.2f}"})\
        .apply(highlight_pastdueinvoice,axis = 1)\
        .set_properties(**{'text-align':'center'})\
        .set_table_attributes('class="table"')\
        .hide_index()\
        .render())

    bootstrap = '<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" crossorigin="anonymous">'
    html_string = bootstrap + html_string

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

    worksheet1.set_column('A:J',11)
    worksheet1.set_column('J:J',15,format) # apply the format to column J, set column width = 18
    worksheet1.set_column('B:B',customer_name_width) # set column A -> J width = longest customer's name
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
    #mail.To = 'accounting@stingerchemicals.com'  #change this line to change receipient's emails
    mail.CC = CCEmails
    mail.Subject = RepFullName + ' Open Invoice as of ' + str(datetime.date.today())
    mail.HTMLBody = '<h2>This is Unpaid Invoices of ' + RepFullName + ' customers</h2>' + html_string

    mail.Attachments.Add(var_output_Excel_path)
    mail.Send()
    os.remove(var_output_Excel_path)       #delete the excel file after email sent

#close ODBC connection
cn.close()






