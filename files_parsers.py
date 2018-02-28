import openpyxl
ONE_PLUS_ONE_FOR_HEADER = 2
OK_COLUMN = 1

def rtv_table(xls_name):
    try:
        xl_workbook = openpyxl.load_workbook(filename=xls_name, read_only=True, data_only=True)
    except FileNotFoundError:
        print(f'Файл {xls_name} не найден')
        return False
    xl_sheet_names = xl_workbook.get_sheet_names()
    xl_sheet = xl_workbook.get_sheet_by_name(xl_sheet_names[0])
    ncols = xl_sheet.max_column
    nrows = xl_sheet.max_row
    columns = {}
    bold_columns = []
    for i in range(1, ncols + 1):
        title = str(xl_sheet.cell(row=1, column=i).value)
        if title:
            columns[title] = i
            if xl_sheet.cell(row=1, column=i).font.bold:
                bold_columns.append(title)
    rows_list = []
    for j in range(2, nrows + 1):
        cur = columns.copy()
        for key, col in columns.items():
            cur[key] = str(xl_sheet.cell(row=j, column=col).value)
        rows_list.append(cur)
    return rows_list, bold_columns


def set_ok(xls_name, row_num_ind):
    while True:
        try:
            xl_workbook = openpyxl.load_workbook(filename=xls_name, read_only=False, data_only=False)
            break
        except FileNotFoundError:
            print(f'Файл {xls_name} не найден')
            return False
        except PermissionError:
            print(f'Файл {xls_name} заблокирован. Сохраните и закойте его. В него будут вноситься отметки об успешности отправки')
            return False
            # continue
    xl_sheet_names = xl_workbook.get_sheet_names()
    xl_sheet = xl_workbook.get_sheet_by_name(xl_sheet_names[0])
    xl_sheet.cell(row=row_num_ind+ONE_PLUS_ONE_FOR_HEADER, column=OK_COLUMN).value = 'ok'
    xl_workbook.save(filename=xls_name)


# import os
# os.chdir(r'C:\Dropbox\repos\batch_email_sender')
# xls_name = 'send_list.xlsx'
# result, bold_columns = rtv_table(xls_name)
# print(bold_columns)
# for row in result:
#     print(row)
