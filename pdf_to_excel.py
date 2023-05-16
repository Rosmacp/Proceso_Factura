import os
import tabula
from openpyxl import Workbook

input_folder = 
output_folder = 
input_file_name = 
output_file_name = 

input_file_path = os.path.join(input_folder, input_file_name)
output_file_path = os.path.join(output_folder, output_file_name)

# Leer la tabla del PDF usando tabula
tables = tabula.read_pdf(input_file_path, pages='all', multiple_tables=True, pandas_options={'header': None})

# Crear un nuevo archivo de Excel
wb = Workbook()
ws = wb.active

# Copiar los datos de la tabla extraída al archivo de Excel
for table in tables:
    for index, row in table.iterrows():
        ws.append(row.tolist())

# Guardar el archivo de Excel
wb.save(output_file_path)
