from my_app.ss_lib.Ssheet_class import Ssheet
from my_app.ss_lib.smartsheet_basic_functions import *
from my_app.func_lib.push_list_to_xls import push_list_to_xls
from my_app.func_lib.push_xls_to_ss import push_xls_to_ss
from my_app.func_lib.push_list_to_ss import push_list_to_ss
from my_app.sheet_map import sheet_map
import time

new_sheet = Ssheet('test fri 13th')
my_list = [['jim', 'stan', 'ang', 'blanche'], ['chris', 'ang', 'amanda']]

if new_sheet.id == -1:
    new_sheet.create_sheet('test fri 13th', sheet_map)
    print('Created', new_sheet.id)
else:
    print('Found', new_sheet.id)


push_list_to_ss(my_list, new_sheet)
# time.sleep(2)
# new_sheet.delete_sheet()
print(new_sheet.id)

exit()



push_list_to_ss(my_list,new_sheet)

# ws_name = 'Tetration Customer Adoption Workspace'
# ws_id = ss_get_ws(new_sheet.ss, ws_name)

# print(ss_get_sheet(new_sheet.ss, 'Tetration Adoption Reference Data')['id'])
# print('workspace', ws_id)
# exit()
# folder_name = 'Tetration Adoption Reference Data'
# print('folders', ss_get_folder_id(new_sheet.ss, folder_name ))
#print(ws_id)





# Get the SmartSheets we need
raw_emails = Ssheet('saas_activation_raw_emails')

raw_rows = raw_emails.get_rows()
scrubbed_list = []
my_row = []

# Grab the headers
for my_col_header in sheet_map:
    my_row.append(my_col_header['title'])
scrubbed_list.append(my_row)

# Loop over the raw rows
for row in raw_rows.items():
    raw_body = row[1]['raw_body']
    my_row = []

    for my_col in sheet_map:
        search_tag = my_col['title']
        search_value = ''
        search_char = ''
        search_pos = raw_body.find(search_tag) + len(search_tag)+2  # Add two for the ":ord(10)" newline char

        if raw_body[search_pos-1] == '=':
            search_pos = search_pos+1

        for search_char in raw_body[search_pos:]:
            if ord(search_char) == 10:
                break
            else:
                search_value = search_value + search_char

        # Clean up the email address
        if search_tag.lower().find('email') != -1:
            search_value = (search_value[0:search_value.find('<')])

        # Capitalize First and Last Names
        if search_tag.find('First Name') != -1 or search_tag.find('Last Name') != -1:
            search_value = search_value.capitalize()

        print(search_tag, search_value)
        my_row.append(search_value)

    scrubbed_list.append(my_row)
    print()

push_list_to_xls(scrubbed_list, 'saas_activation_data.xlsx')

# jim = Ssheet('saas_activation_data', sheet_map)
# jim.create_sheet('saas_activation_data', sheet_map)
# print(jim.id)
#
# push_xls_to_ss('saas_activation_data.xlsx', 'saas_activation_data')






