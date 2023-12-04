#!/usr/bin/env python
# coding: utf-8

# In[23]:


import sys
import time
import pandas as pd
###################################################
####### EXAMPLE ROWS 4 9 10 11 12 13 14      ######
###################################################
#C:\Users\Lazaros.Sofikitis\Desktop\ΔΕΕΔ_ΜΣ_ΑΤΤΙΚΗ (1).xlsx
#file_path = 'C:/Users/Lazaros.Sofikitis/Desktop/ΔΕΕΔ_ΜΣ_ΑΤΤΙΚΗ (1).xlsx'# path to file

output_file = 'combined_sheets.xlsx'
file_path = input("Enter the file path: ")

if len(file_path) >= 2 and file_path[0] == file_path[-1]:
    new_file_path = file_path[1:-1]
else:
    new_file_path = file_path
  # Outputs the file path without double(double quotes!!!) very important to read the filepath

xls = pd.ExcelFile(new_file_path)
sheet_names = xls.sheet_names[1:]

# Get user input for rows to extract

keep_rows = list(map(int, input("Enter row indices to keep (separated by space, input -1 to fetch all data) : ").split()))
print('working on it...')
# Wait for 3 seconds
time.sleep(1)


all_non_negative = all(row_index >= 0 for row_index in keep_rows)
if all_non_negative:
    print('retrieving the specified rows...')
    keep_rows_minus_one = [x - 1 for x in keep_rows]  # pandas counts from 0
    keep_rows_minus_two = [x - 1 for x in keep_rows_minus_one]  # pandas counts from 0
    
elif len(keep_rows) == 1 and -1 in keep_rows:
    print('retrieving all data in the excel sheets...')
    second_sheet_df = pd.read_excel(new_file_path, sheet_name=1)  # Assuming data is in the second sheet
    start_range_1 = 0  # Replace with your start index
    end_range_1 = len(second_sheet_df)   # Replace with your end index        
    first_occurrences = {value: True for value in ['Φυσικοχημική Ανάλυση Ελαίου',
                                                   'Αεριοχρωματογραφική Ανάλυση Διαλυμένων Αερίων', 
                                                   'ΙΔΙΟΤΗΤΑ', 
                                                   'ΠΕΡΙΕΚΤΙΚΟΤΗΤΕΣ ΑΕΡΙΩΝ', 
                                                   'Υδρογόνο (H2)', 
                                                   'Λόγοι ROGERS', 
                                                   'R1 = CH4/H2']}
    indices = []
    values = []
        
    for i in range(start_range_1, end_range_1):
        cell_value = second_sheet_df.iloc[i, 0]  # Access the cell in the first column

        if pd.notnull(cell_value) and (cell_value not in first_occurrences.keys() or first_occurrences[cell_value]):
            indices.append(i)
            values.append(cell_value)

            if cell_value in first_occurrences.keys():
                first_occurrences[cell_value] = False

    result_df = pd.DataFrame({'Index': indices, 'Value': values})
    values_to_remove = [1, 5, 27, 33, 48]

    # Deleting rows based on specified values in 'Índex' column
    result_df = result_df[~result_df['Index'].isin(values_to_remove)]
    result_df = result_df.reset_index(drop=True)
    keep_rows = result_df.iloc[:, 0].tolist()
    keep_rows_minus_one = [x - 1 for x in keep_rows]  # pandas counts from 0
    keep_rows_minus_two = [x - 1 for x in keep_rows_minus_one]  # pandas counts from 0
    
else :
    print('conflict with the input rows detected... shutting down')
    time.sleep(2)
    sys.exit() 

# Parse the second sheet of the Excel file
second_sheet_df = xls.parse(1)  # Index 1 refers to the second sheet (0-indexed)

# Extract values from the first column based on user-provided rows
selected_values = second_sheet_df.iloc[keep_rows_minus_two, 0].tolist()
selected_values2 = second_sheet_df.iloc[keep_rows_minus_two, 1].tolist()
selected_values2.insert(0, 'Name')
# Prepend 'Τοποθεσία' to the list of selected values
selected_values.insert(0, 'Τοποθεσία')
selected_values = [str(item) for item in selected_values]
selected_values2 = [str(item) for item in selected_values2]

#print("Selected Values from the First Column of the Second Sheet:")
delimiter=' '
selected_values3 = [delimiter.join(pair) for pair in zip(selected_values, selected_values2)]
print(' The values you selected with their corresponding units:  ')
print(selected_values3)
# Initialize a list to store data from each sheet
data_frames = []

for idx, sheet_name in enumerate(sheet_names):
    df = xls.parse(sheet_name, header=None)
    df.drop([0, 1], axis=1, inplace=True)
    df = df.iloc[keep_rows_minus_one]
    df = df.transpose()
    df.insert(0, 'Sheet Name', sheet_name)  # Insert sheet name as a column
    data_frames.append(df)

# Concatenate all data frames along rows
combined_data = pd.concat(data_frames, axis=0, ignore_index=True)
combined_data.columns = selected_values3

# Remove rows where all columns except the first have null values
cleaned_df = combined_data.dropna(subset=combined_data.columns[1:], how='all')


cleaned_df.to_excel(output_file, index=False)
print('finished succesfully!')
time.sleep(2)





