# %%
import xlwings as xw
import pandas as pd

# %%
recl_path = r"C:\Users\nil.monreal\Desktop\Nil\EZ Notes\Python Scripts\Pyhton Scripts\AR Automation\AR Reports\2.- RECL JULY 2024.xlsx"
ar51_path = r"C:\Users\nil.monreal\Desktop\Nil\EZ Notes\Python Scripts\Pyhton Scripts\AR Automation\Closing Reports Ry4\AR 5 1 DISTRIBUTION REPORT AUGUST 2024.xlsx"

# %%
# Step 2: Open the first file (recl_path) and the AR 5 1 sheet
recl_wb = xw.Book(recl_path, update_links=False)  # Open the Recl workbook
recl_sheet = recl_wb.sheets['AR 5 1']  # Access the 'AR 5 1' sheet

# %%
# Step 3: Remove any existing filters (reset AutoFilters)
if recl_sheet.api.AutoFilterMode:
    recl_sheet.api.AutoFilterMode = False

# %%
# Step 4: Delete all content from the sheet (clear all data)
recl_sheet.clear()  # This clears the entire sheet's content

# %%
# Step 5: Open the second file (ar51_path) and copy content from its AR 5 1 sheet
ar51_wb = xw.Book(ar51_path, update_links=False)  # Open the AR 5 1 workbook
ar51_sheet = ar51_wb.sheets['AR 5 1']  # Access the 'AR 5 1' sheet

# %%
# Copy all data from AR 5 1 sheet in ar51 file
ar51_data = ar51_sheet.used_range.value  # Get the entire data from the second file

# %%
# Paste the data into the AR 5 1 sheet of the recl file
recl_sheet.range('A1').value = ar51_data  # Paste the data starting at A1

# %%
ar51_wb.close()

# %%
last_row = recl_sheet.range('A' + str(recl_sheet.cells.last_cell.row)).end('up').row

recl_sheet.range(f"A6:L{last_row}").api.AutoFilter(Field=12, Criteria1=['CDZ', 'OCT', 'SAF'], Operator=7)

# recl_sheet.range("A6:L6000").api.AutoFilter(Field=12, Criteria1=['CDZ', 'OCT', 'SAF'], Operator=7)

# %%
# Initialize a list to store the filtered rows that start with '341'
filtered_rows = []

# Iterate through each visible row and check if column E starts with '341'
for i in range(7, 6001):  # Adjust range if necessary
    if recl_sheet.api.Rows(i).Hidden == False:
        # Read the value of column E in this row (5th column in 1-indexed Excel)
        col_e_value = recl_sheet.range(f'E{i}').value
        
        # Check if column E starts with '341'
        if str(col_e_value).startswith('341'):
            # Add the entire row (A:L) to the filtered_rows list
            row_data = recl_sheet.range(f'A{i}:L{i}').value
            formatted_row = []
            
            for item in row_data:
                # If item is a float but is an integer number, convert it to int
                if isinstance(item, float) and item.is_integer():
                    formatted_row.append(int(item))
                else:
                    formatted_row.append(item)
                    
            filtered_rows.append(formatted_row)

# Define headers from row 6 (A6:L6)
headers = recl_sheet.range("A6:L6").value

# Create a DataFrame from the filtered rows
df341 = pd.DataFrame(filtered_rows, columns=headers)

# Convert NaN values to zero
df341.fillna(0, inplace=True)

df341

# %%
df341 = df341[df341['Debit-amt'] != 0]

df341

# %%
row_count = df341.shape[0]
row_count

# %%
try:# Assuming recl_wb is already opened and recl_sheet2 is defined
    recl_sheet2 = recl_wb.sheets["RECL VENTAS-PAGOS"]

    # Start searching from cell C4
    current_row = 2  # Starting row
    last_non_empty_row = None  # To store the last non-empty row

    # Loop down column C starting from C4 until we hit the first blank cell
    while True:
        cell_value = recl_sheet2.range(f'C{current_row}').value

        # Stop if we hit a blank cell
        if not cell_value:
            break

        # If the cell is not empty, store the current row as the last non-empty row
        last_non_empty_row = current_row

        # Move to the next row
        current_row += 1

    recl_sheet2.range(f"B2:J{last_non_empty_row}").clear()
    if row_count > last_non_empty_row:
        recl_sheet2.range(f'{last_non_empty_row}:{row_count+1}').api.Insert()
except:
    pass

# %%
df341 = df341.drop(['GL-code', 'Description', 'Entity'], axis=1)

df341

# %%
df341[['Debit-amt', 'Credit-amt']] = df341[['Credit-amt', 'Debit-amt']]
df341

# %%
#Copying values and setting the sum()
recl_sheet2.range('B2').value = df341.values
recl_sheet2.range(f'D{row_count+4}').formula = f"=sum(D{row_count+1}:D2)"
recl_sheet2.range(f'E{row_count+4}').formula = f"=sum(E{row_count+1}:E2)"

# %%
# Start searching from the given last_non_empty_row
start_row = last_non_empty_row + 1  # You want to start searching after the last non-empty row
column = 'D'  # Column D

# Loop down column D starting from the given row until we find a non-empty cell
current_row = start_row
while True:
    cell_value = recl_sheet.range(f'{column}{current_row}').value

    # If we find a non-empty cell, we stop the loop
    if cell_value:
        break

    # If the cell is empty, clear the entire row
    recl_sheet.range(f'{current_row}:{current_row}').delete()

    # Move to the next row
    current_row += 1

# %%
# Initialize a list to store the filtered rows that do NOT start with '341'
filtered_rows = []

# Iterate through each visible row and check if column E does NOT start with '341'
for i in range(7, 6001):  # Adjust range if necessary
    if recl_sheet.api.Rows(i).Hidden == False:
        # Read the value of column E in this row (5th column in 1-indexed Excel)
        col_e_value = recl_sheet.range(f'E{i}').value
        
        # Check if column E does NOT start with '341'
        if not str(col_e_value).startswith('341'):
            # Add the entire row (A:L) to the filtered_rows list
            row_data = recl_sheet.range(f'A{i}:L{i}').value
            formatted_row = []
            
            for item in row_data:
                # If item is a float but is an integer number, convert it to int
                if isinstance(item, float) and item.is_integer():
                    formatted_row.append(int(item))
                else:
                    formatted_row.append(item)
                    
            filtered_rows.append(formatted_row)

# Define headers from row 6 (A6:L6)
headers = recl_sheet.range("A6:L6").value

# Create a DataFrame from the filtered rows
dfrest = pd.DataFrame(filtered_rows, columns=headers)

# Convert NaN values to zero
dfrest.fillna(0, inplace=True)

dfrest

# %%
dfrest = dfrest[dfrest['Document'] != 0]
dfrest = dfrest[dfrest['Debit-amt'] != 0]

dfrest

# %%
dfrest = dfrest.drop(['GL-code', 'Description', 'Entity'], axis=1)
dfrest['Tran-date'] = pd.to_datetime(dfrest['Tran-date']).dt.strftime('%d/%m/%Y')

df_filtered = dfrest[dfrest['Document'].str.startswith('CM')]

df_filtered[['Debit-amt', 'Credit-amt']] = df_filtered[['Credit-amt', 'Debit-amt']]

rows_to_drop = df_filtered.shape[0]

dfrest = dfrest.iloc[:-rows_to_drop] 

dfrest = pd.concat([dfrest, df_filtered], ignore_index=True)

dfrest

# %%
row_count2 = dfrest.shape[0]
row_count2


# %%
print(row_count + 7)

# %%
# Assuming recl_wb is already opened and recl_sheet2 is defined
# Start searching from first cell
current_row = row_count + 7  # Starting row
starting_row = current_row
last_non_empty_row = None  # To store the last non-empty row

# Loop down column C starting from C4 until we hit the first blank cell
while True:
    cell_value = recl_sheet2.range(f'C{current_row}').value

    # Stop if we hit a blank cell
    if not cell_value:
        break

    # If the cell is not empty, store the current row as the last non-empty row
    last_non_empty_row = current_row

    # Move to the next row
    current_row += 1

recl_sheet2.range(f"B{starting_row}:J{last_non_empty_row}").clear()
if row_count2 > last_non_empty_row:
    recl_sheet2.range(f'{last_non_empty_row + 1}:{last_non_empty_row + (row_count2 - last_non_empty_row)}').api.Insert()

recl_wb.save()

print("Script runned succesfully")





# %%
last_non_empty_row

# %%
recl_sheet2.range(f'B{starting_row}').value = dfrest.values


# %%
# Start searching from the given last_non_empty_row
start_row = last_non_empty_row + 1  # You want to start searching after the last non-empty row
column = 'D'  # Column D

# Loop down column D starting from the given row until we find a non-empty cell
current_row = start_row
while True:
    cell_value = recl_sheet.range(f'{column}{current_row}').value

    # If we find a non-empty cell, we stop the loop
    if cell_value:
        break

    # If the cell is empty, clear the entire row
    recl_sheet.range(f'{current_row}:{current_row}').delete()

    # Move to the next row
    current_row += 1

# %%
current_row

# %%
recl_sheet2.range(f'D{current_row+4}').formula = f"=sum(D{current_row+1}:D{starting_row})"
recl_sheet2.range(f'E{current_row+4}').formula = f"=sum(E{current_row+1}:E{starting_row})"


