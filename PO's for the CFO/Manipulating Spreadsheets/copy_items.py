import pandas as pd
from datetime import datetime

file_path = r"C:\Users\nil.monreal\Desktop\Nil\EZ Notes\Python_Scripts\Pyhton_Scripts\PO's for the CFO\Manipulating Spreadsheets\Spreadsheet\PO's.xlsx"

excel_data = pd.read_excel(file_path, sheet_name=None)

# Get the specific sheets to modify
main_sheet = excel_data['Purchase Orders']
po_sheet = excel_data["PO's"]

# Ensure that specific columns are treated as strings
main_sheet['Purchase Order'] = main_sheet['Purchase Order'].astype(str)
po_sheet['PO-No'] = po_sheet['PO-No'].astype(str)
po_sheet['Item-Number'] = po_sheet['Item-Number'].astype(str)
po_sheet['Rev-Pr-Dt'] = pd.to_datetime(po_sheet['Rev-Pr-Dt'], errors='coerce').dt.strftime('%m/%d/%Y')

# Insert the 'PN' column between 'Purchase Order' and 'Vendor'
if 'PN' not in main_sheet.columns:
    # Find the index of the 'Purchase Order' column
    po_index = main_sheet.columns.get_loc('Purchase Order')
    
    # Insert a new column 'PN' after 'Purchase Order'
    main_sheet.insert(po_index + 1, 'PN', '')

#Insert all the needed columns
if 'Rev-Pr-Dt' not in main_sheet.columns:
    # Find the index of the 'Created on' column
    po_index = main_sheet.columns.get_loc('Created On')
    
    # Insert a new columns
    main_sheet.insert(po_index + 1, 'Rev-Pr-Dt', '')
    main_sheet.insert(po_index + 2, 'Quantity', '')
    main_sheet.insert(po_index + 3, 'Price', '')
    main_sheet.insert(po_index + 4, 'Total', '')
    main_sheet.insert(po_index + 5, 'Cst 54', '')
    main_sheet.insert(po_index + 6, 'Std Ext', '')
    main_sheet.insert(po_index + 7, 'PPV (Fav)', '')
    main_sheet.insert(po_index + 8, 'Comments compras', '')
    main_sheet.insert(po_index + 9, 'Comments Finanzas', '')

# Create a new DataFrame to store the final data
final_sheet = pd.DataFrame(columns=main_sheet.columns)

# Iterate over each row in the main_sheet
for i in range(len(main_sheet)):
    original_row = main_sheet.iloc[i].copy()  # Copy the original row
    final_sheet = pd.concat([final_sheet, pd.DataFrame([original_row])], ignore_index=True)  # Add original row to final sheet
    
    purchase_order = str(original_row['Purchase Order'])
    
    if pd.notna(purchase_order):
        # Find all matching rows in the po_sheet for the current purchase order
        matching_rows = po_sheet[po_sheet['PO-No'] == purchase_order]
        
        if not matching_rows.empty:
            for _, matching_row in matching_rows.iterrows():
                # Create a new blank row based on the original row structure
                new_row = pd.Series(index=main_sheet.columns)
                
                # Populate the new row with the relevant information
                new_row['Purchase Order'] = purchase_order
                new_row['PN'] = matching_row['Item-Number']
                new_row['Quantity'] = matching_row['Ord-Q']
                new_row['Price'] = matching_row['Unit-$']
                new_row['Rev-Pr-Dt'] = matching_row['Rev-Pr-Dt']
                
                # Append the new row below the original row
                final_sheet = pd.concat([final_sheet, pd.DataFrame([new_row])], ignore_index=True)
        else:
            # If no matching rows are found, insert a completely blank row
            blank_row = pd.Series(index=main_sheet.columns)
            final_sheet = pd.concat([final_sheet, pd.DataFrame([blank_row])], ignore_index=True)

# Update the sheet in the dictionary
excel_data['Purchase Orders'] = final_sheet

# Get the current date in MM-DD-YYYY format
current_date = datetime.now().strftime("%m-%d-%Y")

# Define the new file path with the current date in the sheet name
new_file_path = r"C:\Users\nil.monreal\Desktop\Nil\EZ Notes\Python_Scripts\Pyhton_Scripts\PO's for the CFO\Manipulating Spreadsheets\Spreadsheet\PO's.xlsx" + current_date + ".xlsx"

# Write back all sheets to the new Excel file
try:
    with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
        for sheet_name, df in excel_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"File successfully saved as {new_file_path}")
except Exception as e:
    print(f"Unsuccessfully saved error: {e}")



#     Let's optimize the script by addressing performance and readability improvements. The main idea is to avoid frequent concatenations and make the code more efficient.

# Optimized Script

# import pandas as pd
# from datetime import datetime

# # Load the Excel file
# folder_path = input("Paste the path to your file: ")
# file_path = folder_path + r"\PO's.xlsx"

# excel_data = pd.read_excel(file_path, sheet_name=None)

# # Get the specific sheets to modify
# main_sheet = excel_data['Purchase Orders']
# po_sheet = excel_data["PO's"]

# # Ensure that specific columns are treated as strings
# main_sheet['Purchase Order'] = main_sheet['Purchase Order'].astype(str)
# po_sheet['PO-No'] = po_sheet['PO-No'].astype(str)
# po_sheet['Item-Number'] = po_sheet['Item-Number'].astype(str)

# # Insert the 'PN' column between 'Purchase Order' and 'Vendor'
# if 'PN' not in main_sheet.columns:
#     po_index = main_sheet.columns.get_loc('Purchase Order')
#     main_sheet.insert(po_index + 1, 'PN', '')

# # Insert additional columns if 'Rev-Pr-Dt' does not exist
# if 'Rev-Pr-Dt' not in main_sheet.columns:
#     po_index = main_sheet.columns.get_loc('Created On')
#     additional_columns = ['Rev-Pr-Dt', 'Quantity', 'Price', 'Total', 'Cst 54', 'Std Ext', 'PPV (Fav)', 'Comments compras', 'Comments Finanzas']
#     for idx, col in enumerate(additional_columns):
#         main_sheet.insert(po_index + 1 + idx, col, '')

# # Initialize a list to hold all rows for final DataFrame
# final_rows = []

# # Iterate over each row in the main_sheet
# for i in range(len(main_sheet)):
#     original_row = main_sheet.iloc[i].copy()
#     final_rows.append(original_row)  # Add original row to final sheet

#     purchase_order = original_row['Purchase Order']

#     if pd.notna(purchase_order):
#         # Find all matching rows in the po_sheet for the current purchase order
#         matching_rows = po_sheet[po_sheet['PO-No'] == purchase_order]

#         if not matching_rows.empty:
#             for _, matching_row in matching_rows.iterrows():
#                 # Create a new row and populate it with relevant information
#                 new_row = pd.Series(index=main_sheet.columns)
#                 new_row['Purchase Order'] = purchase_order
#                 new_row['PN'] = matching_row['Item-Number']
#                 new_row['Quantity'] = matching_row['Ord-Q']
#                 new_row['Price'] = matching_row['Unit-$']
#                 final_rows.append(new_row)  # Add the new row to final list
#         else:
#             # If no matching rows are found, append a blank row
#             blank_row = pd.Series(index=main_sheet.columns)
#             final_rows.append(blank_row)

# # Convert the final_rows list to a DataFrame
# final_sheet = pd.DataFrame(final_rows)

# # Update the sheet in the dictionary
# excel_data['Purchase Orders'] = final_sheet

# # Get the current date in MM-DD-YYYY format
# current_date = datetime.now().strftime("%m-%d-%Y")

# # Define the new file path with the current date in the sheet name
# new_file_path = folder_path + r"\PO's_" + current_date + ".xlsx"

# # Write back all sheets to the new Excel file
# try:
#     with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
#         for sheet_name, df in excel_data.items():
#             df.to_excel(writer, sheet_name=sheet_name, index=False)
#     print(f"File successfully saved as {new_file_path}")
# except Exception as e:
#     print(f"Unsuccessfully saved error: {e}")

# Key Optimizations:
# Avoiding Frequent Concatenations:

# Instead of using pd.concat in every iteration, we store rows in a list (final_rows). At the end, we convert this list to a DataFrame, which is much faster and avoids the overhead of repeated concatenations.
# Predefining Column Data Types:

# Ensuring that columns are cast to str type at the beginning helps avoid data type inconsistencies that could cause warnings or errors.
# Reduced Redundancy:

# Removed redundant checks and processes by simplifying the control flow.
# Simplified Column Insertion:

# For inserting multiple columns at once, I used a loop to reduce the repetition in the code.
# This version of the script should be more efficient, especially for large datasets, and should run faster while consuming less memory. If you notice any further opportunities for optimization or encounter issues, feel free to let me know!


