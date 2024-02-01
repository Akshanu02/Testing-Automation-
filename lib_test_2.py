import subprocess
import openpyxl

# Specify the path to the application executable
application_path = '/etc/alternatives/gnome-text-editor'  # Example path for the calculator on Linux

# Specify the path to the Excel file
excel_file_path = 'automation_library.xlsx'

try:
    # Open the application
    subprocess.Popen([application_path])
    print(f"Application {application_path} opened successfully.")

    # Read data from Excel file
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active
    column_values = [cell.value for cell in sheet['A']]
    
    # Convert cell value to string and write it into the application
    for cell_value in column_values:
        subprocess.run([application_path, str(cell_value)])
        print(f"Data from Excel written into the application: {cell_value}")
except Exception as e:
    print(f"Error: {e}")

