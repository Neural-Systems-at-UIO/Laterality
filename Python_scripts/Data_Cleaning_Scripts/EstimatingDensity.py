import pandas as pd
import numpy as np
import math
import os

# Function to prompt the user for a file and read it into a dataframe
def read_dataframe(prompt):
    file_path = input(prompt)
    return pd.read_excel(file_path, index_col=0)

# Function to prompt the user for the output file path
def get_output_path(prompt):
    output_path = input(prompt)
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    return output_path

# Prompt the user for the three Excel files and the rat number
counts_df = read_dataframe("Please enter the path to the object counts Excel file: ")
region_areas_df = read_dataframe("Please enter the path to the region areas Excel file: ")
object_sizes_df = read_dataframe("Please enter the path to the object sizes Excel file: ")
rat_number = str(input("Please enter the rat number: "))

# Add a 'Total' row to the region areas dataframe
region_areas_df.loc['Total'] = region_areas_df.sum()

# Create a new dataframe with the same structure as object sizes, initialized with NaNs
new_df = pd.DataFrame(index=object_sizes_df.index, columns=object_sizes_df.columns)
new_df[:] = np.nan

# Fill the new dataframe based on the calculation rule provided
for column in new_df.columns:
    column_index = new_df.columns.get_loc(column)
    for i, row in enumerate(new_df.index):
        counts_value = counts_df.iloc[i, column_index]  # Use iloc for integer location
        area_value = region_areas_df.iloc[i, column_index]  # Match the corresponding area value
        size_value = object_sizes_df.iloc[i, column_index]
        
        # Calculate the new value
        if pd.notnull(counts_value) and pd.notnull(area_value) and pd.notnull(size_value):
            if area_value != 0:
                new_value = counts_value / (area_value * 0.1936 * (10**(-9)) * (40 + 2 * math.sqrt(size_value / math.pi)))
                new_df.iloc[i, column_index] = new_value

# Prompt the user for the output directory and file name
output_directory = get_output_path("Please enter the output directory for the new Excel file: ")
file_name = f'{rat_number}_densities.xlsx'
output_path = os.path.join(output_directory, file_name)

# Save the new dataframe to an Excel file
new_df.to_excel(output_path, index=True)

print(f"New dataframe saved to {output_path}")
