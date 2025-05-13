from Utils.copy_excel_range import copy_excel_range


def footer_format_copy(source_wb, destiny_wb, source_sheet_name, last_row_to):
    

    footer_range = "A1:AQ4"
    last_row = last_row_to + 5
    
    destiny_ws = destiny_wb["Reporte"]
    
    #try:
        
        #copy_excel_range(source_sheet_name, destiny_ws, )