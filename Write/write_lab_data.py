from Utils.merged_cell import merged_cell


def write_lab_data(wb, data: list, start_row, client_sample_id):
    sheet_to_write = wb["Reporte"]
    start_row = start_row + 2

    print(f"START ROW PARA LAB DATA {start_row}")

    try:
        for index, row in enumerate(data):
            important_data = row[0]


            if len(important_data) > 0:
                # Generar el ID consecutivo
                formatted_id = f"{client_sample_id}-{index + 1:03d}"

                # Extraer los datos
                item = important_data[0]
                lab_sample_id = important_data[1]
                collected = important_data[2]
                collected_time = important_data[3]
                sample_matrix = important_data[5]
                analysis_requested = "PENDIENTE"

                # Escribir cada dato en su columna correspondiente
                sheet_to_write[f"G{start_row}"].value = formatted_id  # ID consecutivo
                sheet_to_write[f"B{start_row}"].value = item
                sheet_to_write[f"K{start_row}"].value = lab_sample_id
                sheet_to_write[f"Q{start_row}"].value = collected
                sheet_to_write[f"U{start_row}"].value = collected_time
                sheet_to_write[f"X{start_row}"].value = sample_matrix
                sheet_to_write[f"AC{start_row}"].value = analysis_requested

                start_row += 1

    except Exception as ex:
        print(f"Error: {ex}")