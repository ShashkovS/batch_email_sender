import openpyxl
import alerts
import os
ONE_PLUS_ONE_FOR_HEADER = 2
OK_COLUMN = 1
OKOK = 'ok'
ORIGINAL_ROW_NUM = 'row_num_WcCRve89'

def rtv_table(xls_name):
    try:
        xl_workbook = openpyxl.load_workbook(filename=xls_name, read_only=True, data_only=True)
    except FileNotFoundError:
        alerts.alert(f'Файл {xls_name} не найден')
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
        cur[ORIGINAL_ROW_NUM] = j
        for key, col in columns.items():
            cur[key] = str(xl_sheet.cell(row=j, column=col).value).replace('None', '')
        rows_list.append(cur)
    return rows_list, bold_columns


def set_ok(xls_name, row_num_real):
    while True:
        try:
            xl_workbook = openpyxl.load_workbook(filename=xls_name, read_only=False, data_only=False)
            break
        except FileNotFoundError:
            alerts.alert(f'Файл {xls_name} не найден')
            return False
        except PermissionError:
            alerts.alert(f'Файл {xls_name} заблокирован. Сохраните и закойте его. В него будут вноситься отметки об успешности отправки')
            return False
            # continue
    xl_sheet_names = xl_workbook.get_sheet_names()
    xl_sheet = xl_workbook.get_sheet_by_name(xl_sheet_names[0])
    xl_sheet.cell(row=row_num_real, column=OK_COLUMN).value = OKOK
    xl_workbook.save(filename=xls_name)


def rtv_template(template_name):
    try:
        with open(template_name, encoding='utf-8') as f:
            template = f.read()
    except FileNotFoundError:
        alerts.alert(f'Файл с шаблоном {template_name} не найден')
        return False
    return template


def rtv_table_and_template(xls_name, template_name):
    rows_list, bold_columns = rtv_table(xls_name)
    template = rtv_template(template_name)
    NOTHING = None, None, None
    if not rows_list:
        alerts.alert(f'В файле {xls_name} не обнаружены строки с данными')
        return NOTHING
    first_data_row = rows_list[0]
    # Проверяем обязательные столбцы: ok, email, subject
    for col_name in ['ok', 'email', 'subject']:
        if col_name not in first_data_row:
            alerts.alert(f'В таблице {xls_name} обязательно должен быть столбец {col_name}')
            return NOTHING
    # Проверяем, что есть всё, что указано в шаблоне
    try:
        template.format(**first_data_row)
    except KeyError as e:
        alerts.alert(f'В таблице {xls_name} должен быть столбец {e!s}, так как он упоминается в шаблоне {template_name}')
        return NOTHING
    # Теперь проверяем существование всех вложений
    attach_cols = [key for key in first_data_row if key.startswith('attach')]
    if attach_cols:
        for rn, row in enumerate(rows_list):
            if '@' not in row['email'] or row['ok'].lower() == OKOK.lower():
                continue
            for attach_key in attach_cols:
                attach_name = row[attach_key]
                if attach_name and not os.path.isfile(attach_name):
                    alerts.alert(f'В таблице {xls_name} в строчке {rn+ONE_PLUS_ONE_FOR_HEADER} в столбце {attach_key} указано вложение "{attach_name}". Этот файл не найден')
                    return NOTHING
    # Проверили, что всё работает. Проверили, что вложения существуют
    return rows_list, bold_columns, template

# import os
# os.chdir(r'C:\Dropbox\repos\batch_email_sender')
# xls_name = 'send_list.xlsx'
# template_name = 'send_template.html'
# rows_list, bold_columns, template = rtv_table_and_template(xls_name, template_name)
# print(rows_list, bold_columns, template)

# result, bold_columns = rtv_table(xls_name)
# alerts.alert(bold_columns)
# for row in result:
#     alerts.alert(row)
