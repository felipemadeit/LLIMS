from os.path import samefile
from Utils.get_wb_sheets import get_wb_sheets
import re


def is_matching_sample(target_id, cell_value):
    # Convertir a string y limpiar espacios
    target_id = str(target_id).strip()
    current_id = str(cell_value).strip()

    # Crear patrón regex que coincida con el ID base y variantes permitidas
    pattern = re.compile(rf'^{re.escape(target_id)}(?:\s+[A-Z]+)?$')
    return bool(pattern.fullmatch(current_id))


def matrix_data_reader(wb_to_read, chain_data):

    matrix_mapping = {
        "Be": "Beryllium",
        "Cd": "Cadmium",
        "Mn": "Manganese",
        "Ag": "Silver",
        "As": "Arsenic",
        "Ba": "Barium",
        "Co": "Cobalt",
        "Cr": "Chromium",
        "Cu": "Copper",
        "Fe": "Iron",
        "Ni": "Nickel",
        "Pb": "Lead",
        "Sb": "Antimony",
        "Se": "Selenium",
        "Sr": "Strontium",
        "Tl": "Thallium",
        "V": "Vanadium",
        "Zn": "Zinc",
        "Al": "Aluminum",
        "Ca": "Calcium",
        "Mg": "Magnesium",
        "K": "Potassium",
        "Na": "Sodium",
        "Hg": "Mercury"
    }

    # Mapeo inverso
    inverse_mapping = {v: k for k, v in matrix_mapping.items()}

    # Recorrer cada fila de datos en chain_data
    for row_idx, row in enumerate(chain_data):
        if not row or not row[0]:  # Si no hay datos principales, saltar
            continue

        sample_id = str(row[0][1]).strip()  # ID de muestra normalizado
        sheet_list = row[1]  # Lista de hojas donde buscar

        # Inicializar lista para almacenar datos de matrices
        if len(row) < 3:  # Si no existe la posición para datos
            row.append([])  # Añadir lista vacía para datos de matrices
        else:
            row[2] = []  # Limpiar datos anteriores

        print(f"\nBuscando muestra: {sample_id}")

        # Buscar en cada hoja especificada para esta muestra
        for matrix_name in sheet_list:
            matrix_name = matrix_name.strip()
            print(f"\nProcesando hoja: {matrix_name}")

            # Determinar el nombre correcto de la hoja
            sheet_to_read = None
            possible_names = [
                matrix_name,  # Primero intentar con el nombre exacto
                inverse_mapping.get(matrix_name),  # Luego con versión corta
                matrix_mapping.get(matrix_name)  # Luego con versión larga
            ]

            # Buscar la primera variación del nombre que exista en el workbook
            for name in possible_names:
                if name and name in wb_to_read.sheetnames:
                    sheet_to_read = wb_to_read[name]
                    break

            if not sheet_to_read:
                print(f"Hoja {matrix_name} no encontrada en el workbook")
                continue

            print(f"Leyendo hoja: {sheet_to_read.title}")

            # Buscar todas las ocurrencias del sample_id en la hoja
            found_any = False
            start_row = 21  # Fila inicial de búsqueda

            while True:
                cell_b = f"B{start_row}"
                cell_value = sheet_to_read[cell_b].value

                if cell_value is None:  # Fin de datos
                    if not found_any:
                        print(f"No se encontraron datos para {sample_id} en esta hoja")
                    break

                # Usar la función de comparación precisa
                if is_matching_sample(sample_id, cell_value):
                    found_any = True
                    print(f"Coincidencia exacta encontrada en fila {start_row}:")
                    print(f"ID en hoja: {cell_value} | ID buscado: {sample_id}")

                    # Recoger datos de las columnas especificadas
                    matrix_data = {
                        'matrix_name': sheet_to_read.title,
                        'row_number': start_row,
                        'data': {
                            'A': sheet_to_read[f"A{start_row}"].value,
                            'B': sheet_to_read[f"B{start_row}"].value,
                            'H': sheet_to_read[f"H{start_row}"].value,
                            'I': sheet_to_read[f"I{start_row}"].value,
                            'J': sheet_to_read[f"J{start_row}"].value
                        }
                    }

                    # Agregar los datos a la fila correspondiente
                    row[2].append(matrix_data)

                    #print(f"Datos recolectados: {matrix_data}")

                start_row += 1
    #for row in chain_data:
        #print(row)
    return chain_data