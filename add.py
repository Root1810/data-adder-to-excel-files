import openpyxl
 
file_path = 'Nonkernel.xlsx'
workbook = openpyxl.load_workbook(file_path)
sheet_names = workbook.sheetnames
 
print("Available sheets:")
for idx, sheet_name in enumerate(sheet_names, start=1):
    print(f"{idx}. {sheet_name}")
 
selected_index = int(input("Select the sheet number you want to edit: ")) - 1
if selected_index < 0 or selected_index >= len(sheet_names):
    print("Invalid selection. Please run the script again and select a valid sheet number.")
else:
    selected_sheet_name = sheet_names[selected_index]
    sheet = workbook[selected_sheet_name]
 
    server_name_col = None
    for cell in sheet[1]:
        if cell.value == "Server Name":
            server_name_col = cell.column
            break
 
    if server_name_col is None:
        print('The "Server Name" column was not found in the selected sheet.')
    else:
        column_letter = openpyxl.utils.get_column_letter(server_name_col)
 
        existing_server_names = set()
        for row in range(2, sheet.max_row + 1):
            cell_value = sheet[f"{column_letter}{row}"].value
            if cell_value:
                existing_server_names.add(cell_value.strip())
 
        txt_file_path = 'data_add.txt'
        with open(txt_file_path, 'r') as file:
            data = file.readlines()
 
        added_any = False
        for line in data:
            hostname = line.strip()
 
            if hostname in existing_server_names:
                print(f"{hostname} already exists.")
            else:
                row_idx = 2
                while sheet[f"{column_letter}{row_idx}"].value:
                    row_idx += 1
                sheet[f"{column_letter}{row_idx}"] = hostname
                existing_server_names.add(hostname)
                added_any = True
 
        if added_any:
            workbook.save(file_path)
            print(f"Data added to the 'Server Name' column ({column_letter}) in sheet {selected_sheet_name} and saved successfully.")
        else:
            print("No new data was added as all entries already exist.")
