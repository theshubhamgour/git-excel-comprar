# The pandas library is used to read in and manipulate the Excel files. The datetime library is used to create a timestamp for the output file name.
import pandas as pd

# Read in both files
file1 = pd.read_excel('file1.xlsx', sheet_name=None)
file2 = pd.read_excel('file2.xlsx', sheet_name=None)

# Loop through each sheet in file2
for sheet_name in file2.keys():
    # Get the corresponding sheets from both files
    sheet1 = file1[sheet_name]
    sheet2 = file2[sheet_name]

    # Find the rows where the values differ between the two sheets
    different_rows = sheet2[~(sheet2 == sheet1)]

    # If there are any different rows, write them to a new file
    if not different_rows.empty:
        with pd.ExcelWriter(f'different_rows_{sheet_name}.xlsx') as writer:
            different_rows.to_excel(writer, sheet_name=sheet_name, index=False)
