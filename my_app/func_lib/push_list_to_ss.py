from my_app.sheet_map import sheet_map


def push_list_to_ss(my_list, ss_obj):
    # This func depends on the sheet_map to provide proper
    # mapping of my_list to the destination SmartSheet (ss_obj)

    print('PUSHING TO >>>>>>>>>> ', ss_obj.name)

    # Need to format my_list as a list of dicts for each row as follows:
    # [[{"strict": false, "columnId": 1, "value": "jim"}], [{"strict": false, "columnId": 2, "value": "stan"}] ]

    my_new_list = []

    for my_row_num, my_row in enumerate(my_list):
        tmp_list = []

        for my_col_num, my_col_val in enumerate(my_row):
            my_new_row_dict = {}

            my_col_title = sheet_map[my_col_num]['title']
            my_col_id = ss_obj.col_name_idx[my_col_title]

            # my_new_row_dict['column_number'] = my_col_num
            # my_new_row_dict['title'] = my_col_title
            my_new_row_dict['strict'] = False
            my_new_row_dict['columnId'] = my_col_id
            my_new_row_dict['value'] = my_col_val

            tmp_list.append(my_new_row_dict)

        my_new_list.append(tmp_list)

    # Push the my_new_list to the destination SmartSheet
    ss_obj.add_rows(my_new_list)

    return
