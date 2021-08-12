import xlsxwriter


def list_to_excel(list_name,headers, file_name):
    first_r = 0
    first_c = 0
    last_r = len(list_name)
    if isinstance(list_name[0],list):
        last_c = len(list_name[0])-1
    else:
        last_c = 0
    workbook = xlsxwriter.Workbook(file_name + '.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.add_table(first_r, first_c, last_r, last_c, {'data': list_name,'columns':headers,
                                                           'style': 'Table Style Light 11'})
    workbook.close()


