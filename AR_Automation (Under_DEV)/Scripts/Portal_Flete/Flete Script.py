# %% [markdown]
# # Script Flete

# %%
import xlwings as xw
import pandas as pd

# %%
path = input("Please add the path to the file...") + r"\EZ Air Shipments Report.xlsx"

try:
    flete_wb = xw.Book(path, update_links=False)  # Open workbook
except Exception as e:
    print(f"Error: {e}")


# %%
act_sheet = flete_wb.sheets['Activity'] 

# %%
last_row = act_sheet.range('A3').end('down').row

last_row

# %%
act_sheet.range(f"A3:AU{last_row}").api.AutoFilter(Field=44, Criteria1="<>", Operator=1)

# %%
# Define your specific range
visible_range = act_sheet.range(f"AU:AU").api.SpecialCells(12)

# %%
values = act_sheet.range(f"AU3:AU{last_row}").value

# Create DataFrame (skip header if needed)
df = pd.DataFrame(values[1:], columns=['AU'])

# %%
df

# %%
def clean_au_value(value):
    if isinstance(value, str):
        # Remove double CH instances
        value = value.replace('CHCH', 'CH')
        # Add CH if missing and keep only digits after
        if not value.startswith('CH'):
            return 'CH' + ''.join(filter(str.isdigit, value))
        return 'CH' + ''.join(filter(str.isdigit, value[2:]))
    elif isinstance(value, (int, float)):
        # Handle numeric values by converting to string and adding CH
        return 'CH' + str(int(value))
    return value

# Apply the cleaning function to your AU column
df['AU'] = df['AU'].apply(clean_au_value)

# %%
df

# %%
# Display unique values to check the patterns
print("Unique values in the column:")
print(df['AU'].unique())

# Check if all values start with 'CH'
print("\nValues not starting with 'CH':")
print(df[~df['AU'].str.startswith('CH')])

# Check the length of each value (should be 8 if format is CHxxxxxx)
print("\nValues not 8 characters long:")
print(df[df['AU'].str.len() != 8])

# Display value counts to see frequency of each value
print("\nValue counts:")
print(df['AU'].value_counts().head())

# Check for any remaining double 'CHCH'
print("\nValues containing 'CHCH':")
print(df[df['AU'].str.contains('CHCH', na=False)])

# %%
# Convert DataFrame column to list of values
cleaned_values = df['AU'].tolist()

for i, value in enumerate(cleaned_values, start=4):
    act_sheet.range(f'AU{i}').value = value

# %%
flete_wb.save()
flete_wb.close()


