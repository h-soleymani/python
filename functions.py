import xlsxwriter
import win32com.client


def list_to_excel(list_name,headers, file_name):
    first_r = 0
    first_c = 0
    last_r = len(list_name)
    if isinstance(list_name[0],list):
        last_c = len(list_name[0])-1
    else:
        last_c = 0
    workbook = xlsxwriter.Workbook('export/' + file_name + '.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.add_table(first_r, first_c, last_r, last_c, {'data': list_name,'columns':headers,
                                                           'style': 'Table Style Light 11'})
    workbook.close()



def send_email (name, recipient, cc,html):
    outlook = win32com.client.Dispatch('outlook.application')
    myhtml = open("html.txt",'r')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.bcc = 'h.soleymani@asiatech.ir'
    mail.Subject = name
    mail.HTMLBody = html
    # mail.Attachments.Add('export/' + name + '.xlsx')
    mail.Attachments.Add('C:/Users/h.soleymani/Desktop/projects/split excel and send email/export/' + name + '.xlsx')
    mail.CC = cc
    mail.Send()



def search (list,name):
   
    for item in list:
        if item[0]==name:
            # print(item[1])
            return item[1]
    return ""    

