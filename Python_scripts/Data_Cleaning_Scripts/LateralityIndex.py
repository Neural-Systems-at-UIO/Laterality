import pandas as pd
import os

# Function to calculate laterality index for a given rat number
def calculate_laterality_index(rat_number, df):
    densities_l = f'{rat_number}_densities L'
    densities_r = f'{rat_number}_densities R'
    df[f'{rat_number}_LI'] = (df[densities_l] - df[densities_r]) / (df[densities_l] + df[densities_r])

# Ask for the file path input
file_path = input("Please enter the path to the Excel file: ")

# Load the data from the Excel file
data = pd.read_excel(file_path)

# Apply the laterality index calculation for each rat number
rat_numbers = ['03', '04', '05', '06']
for rat_number in rat_numbers:
    calculate_laterality_index(rat_number, data)

# Construct the new file name and save the updated Excel file
dir_name, file_name = os.path.split(file_path)
base_name, extension = os.path.splitext(file_name)
new_file_name = f"{base_name}_with_LI{extension}"
new_file_path = os.path.join(dir_name, new_file_name)
data.to_excel(new_file_path, index=False)

print(f"Updated file saved as: {new_file_path}")
