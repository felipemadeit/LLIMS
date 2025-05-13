from Utils.merged_cell import merged_cell


def write_header_data(wb_destiny, header_data):

    sheet_to_write = wb_destiny["Reporte"]

    try:

        company_name= header_data[0]
        client_name = header_data[0]
        client_addres = header_data[1]
        lab_received_date = header_data[9]
        project_location = header_data[8]
        client_phone = header_data[4]




        merged_cell(sheet_to_write, "K7", company_name)
        merged_cell(sheet_to_write, "K8", client_name)
        merged_cell(sheet_to_write, "K9", client_addres)
        merged_cell(sheet_to_write, "AK6", lab_received_date)
        merged_cell(sheet_to_write, "AK8", client_name)
        merged_cell(sheet_to_write, "AK9", client_phone)


    except Exception as ex:

        print(f"ERROR: {ex}")


