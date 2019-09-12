from my_app.ss_lib.Ssheet_class import Ssheet
import time
# Get the SmartSheets we need
raw_emails = Ssheet('saas_activation_raw_emails')
saas_data = Ssheet('saas_activation_data')

# raw_sender_col_id = raw_emails.col_name_idx['sender']
# raw_subject_col_id = raw_emails.col_name_idx['subject']
# raw_body_col_id = raw_emails.col_name_idx['raw_body']
# print(raw_sender_col_id, raw_subject_col_id, raw_body_col_id)


search_tags = ['Bill To ID:', 'Partner Account Name:', 'End Customer Name:']
raw_rows = raw_emails.get_rows()
for row in raw_rows.items():
    raw_body = row[1]['raw_body']

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
    print()

    #print(raw_body.find('Partner Account Name'))
    #print(raw_body.find('End Customer Name'))


    # 'Bill To ID'
    # 'Partner Account Name'
    # 'End Customer Name'




