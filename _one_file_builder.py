# -*- coding: utf-8 -*-
from itertools import chain
import re
import os

DESTINATION_FILE = os.path.join('batch_email_sender', '__batch_email_sender__.py')
module_list = [
    'files_parsers',
    'email_stuff',
    'ui_email_and_passw',
    'ui_main_window',
    'batch_email_sender',
]


file_texts = [open(filename + '.py', 'r', encoding='utf-8').readlines() for filename in module_list]

module_list.append('ensure_modules')

import_lines = [row for row in chain(*file_texts)
                if (row.startswith('import ') or row.startswith('from '))]
import_lines = [imp_row for imp_row in set(import_lines) if all(mod_name not in imp_row for mod_name in module_list)]

joined_text = open('ensure_modules.py', 'r', encoding='utf-8').readlines()
joined_text.append('\n\n# All imports\n')
joined_text.extend(import_lines)
joined_text.append('\n')

for filename, file_text in zip(module_list, file_texts):
    joined_text.append('\n\n###### ' + filename + ' ######\n')
    joined_text.extend(row for row in file_text
                       if not (row.startswith('import ') or row.startswith('from ') or row.startswith('#')))

joined_text = ''.join(joined_text)

replacers = [re.compile(r"\b" + modname + "\.") for modname in module_list]
for regex in replacers:
    joined_text = regex.sub('', joined_text)

with open(DESTINATION_FILE, 'w', encoding='utf-8') as f:
    f.write(joined_text)



import re, requests
re.findall(r'(?s)<tr>' + r'\s*<td>\s*(.*?)\s*</td>' * 7 + r'\s*</tr>',
           requests.post(r'https://reg.olimpiada.ru/rusolymp-summary/search/',
              data=dict(login='sch778179', stage='_', subject='%', year=2017))
              .content.decode('utf-8')
              .replace(r'<!--  -->', '').replace('\n\n', ''))
