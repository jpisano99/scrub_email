from my_app.ss_lib.Ssheet_class import Ssheet
from my_app.func_lib.push_list_to_xls import push_list_to_xls
from my_app.func_lib.push_xls_to_ss import push_xls_to_ss
from my_app.sheet_map import sheet_map
import time


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






