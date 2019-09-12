from my_app.ss_lib.Ssheet_class import Ssheet
from my_app.func_lib.push_list_to_xls import push_list_to_xls
from my_app.func_lib.push_xls_to_ss import push_xls_to_ss

import time
# Get the SmartSheets we need
raw_emails = Ssheet('saas_activation_raw_emails')
saas_data = Ssheet('saas_activation_data')

# raw_sender_col_id = raw_emails.col_name_idx['sender']
# raw_subject_col_id = raw_emails.col_name_idx['subject']
# raw_body_col_id = raw_emails.col_name_idx['raw_body']
# print(raw_sender_col_id, raw_subject_col_id, raw_body_col_id)


search_tags = ['Bill To ID:',
               'Partner Account Name:',
               'End Customer Name:',
               'Subscription Reference ID:',
               'CCW Web Order ID:',
               'Ship To ID:',
               'Order Contact First Name:',
               'Order Contact Last Name:',
               'Order Contact Email:',
               'Order Contact Phone:',
               'Billing Contact First Name:',
               'Billing Contact Last Name:',
               'Billing Contact Email:',
               'Billing Contact Phone:',
               'Reseller Contact First Name:',
               'Reseller Contact Last Name:',
               'Reseller Contact Email:',
               'Reseller Contact Phone:',
               'Service To Contact First Name:',
               'Service To Contact Last Name:',
               'Service To Contact Phone:',
               'Initial Term:',
               'Renewal Term:',
               'Prepayment Term:',
               'Billing Model:',
               'MRR:',
               'Route to Market:',
               'Service Requested Start Date:']

raw_rows = raw_emails.get_rows()
scrubbed_list = []
my_row = []
for my_col_header in search_tags:
    my_row.append(my_col_header[:-1])
scrubbed_list.append(my_row)

for row in raw_rows.items():
    raw_body = row[1]['raw_body']
    my_row = []

    for search_tag in search_tags:
        search_value = ''
        search_char = ''
        search_pos = raw_body.find(search_tag) + len(search_tag)+1  # Add one for the ord(10) newline char

        for search_char in raw_body[search_pos:]:
            if ord(search_char) == 10 :
                break
            else:
                search_value = search_value + search_char
        print(search_tag, search_value)
        my_row.append(search_value)

    scrubbed_list.append(my_row)
    print()

push_list_to_xls(scrubbed_list, 'junk.xlsx')
push_xls_to_ss('junk.xlsx', 'saas_activation_data')






