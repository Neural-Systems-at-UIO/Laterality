import pandas as pd
import os

# Function to convert slice number to the desired format
def convert_to_qc_format(slice_number):
    return f's{slice_number:03}'

def region_mapping(column_name, qc_extended_df):
    # Step 1: Extract slice number and region names from non-null values in column
    # Create a dictionary to store slice number as key and list of region names as value
    region_mapping = {}

    for index, row in column_name.items():
        # Check if the row is a string
        if isinstance(row, str):
            slice_number = qc_extended_df.iloc[index, 0]  # Extract slice number
            regions = [region.strip() for region in row.split(';')]  # Split by semicolon and strip whitespace
            region_mapping[slice_number] = regions

    return region_mapping

def highlight_cells_modified(row):
    return ['background-color: yellow' if (row.name, col) in modifications_objects else '' for col in row.index]

# Function to generate the new filename with '_adjustments' appended
def adjusted_filename(file_path):
    folder_path, file_name = os.path.split(file_path)
    base_name, ext = os.path.splitext(file_name)
    new_name = f"{base_name}_adjustments{ext}"
    return os.path.join(folder_path, new_name)

objects_path = input("Please enter the path to the objects file: ")
objects_df = pd.read_excel(objects_path)
regions_path = input("Please enter the path to the regions file: ")
regions_df = pd.read_excel(regions_path)
qc_extended_path = input("Please enter the path to the quality control (qc) file: ")
qc_extended_df = pd.read_excel(qc_extended_path)

# Convert the labels in 03_objects.xlsx and 03_regions.xlsx
objects_df['Unnamed: 0'] = objects_df['Unnamed: 0'].apply(convert_to_qc_format)
regions_df['Unnamed: 0'] = regions_df['Unnamed: 0'].apply(convert_to_qc_format)

# Handling parts adjustments
parts_adjustment = qc_extended_df["parts adjustment"]
non_null_values_parts_adjustment = parts_adjustment.dropna()
num_non_null_values_parts_adjustment = len(non_null_values_parts_adjustment)
if num_non_null_values_parts_adjustment > 0:
    colliculi_objects_path = input("Please enter the path to the colliculi cell counts table: ")
    colliculi_regions_path = input("Please enter the path to the colliculi region pixels table: ")
        
    colliculi_objects_df = pd.read_excel(colliculi_objects_path)
    colliculi_regions_df = pd.read_excel(colliculi_regions_path)
    colliculi_objects_df['Unnamed: 0'] = colliculi_objects_df['Unnamed: 0'].apply(convert_to_qc_format)
    colliculi_regions_df['Unnamed: 0'] = colliculi_regions_df['Unnamed: 0'].apply(convert_to_qc_format)

    region_mapping = region_mapping(parts_adjustment, qc_extended_df)
    
    # Step 2: Locate the corresponding cell in 03_objects.xlsx and 03_regions.xlsx
    # Dictionary to store the extracted values and their coordinates
    extracted_values_objects = {}
    extracted_values_regions = {}

    for slice_number, regions in region_mapping.items():
        for region in regions:
            if region in objects_df.columns:  # Ensure that the region exists as a column in the dataframe
                value_objects = objects_df.loc[objects_df['Unnamed: 0'] == slice_number, region].values[0]
                value_regions = regions_df.loc[regions_df['Unnamed: 0'] == slice_number, region].values[0]
                
                # Store the values and their coordinates
                extracted_values_objects[(slice_number, region)] = value_objects
                extracted_values_regions[(slice_number, region)] = value_regions

    # Step 3: Replace the cells in 03_objects.xlsx and 03_regions.xlsx with the corresponding cells from the new files
    modifications_objects = {}
    modifications_regions = {}

    for (slice_number, region), _ in extracted_values_objects.items():
        if region in colliculi_objects_df.columns:  # Ensure that the region exists as a column in the dataframe
            new_value_objects = colliculi_objects_df.loc[colliculi_objects_df['Unnamed: 0'] == slice_number, region].values[0]
            new_value_regions = colliculi_regions_df.loc[colliculi_regions_df['Unnamed: 0'] == slice_number, region].values[0]
            
            # Store the modifications
            modifications_objects[(slice_number, region)] = (extracted_values_objects[(slice_number, region)], new_value_objects)
            modifications_regions[(slice_number, region)] = (extracted_values_regions[(slice_number, region)], new_value_regions)
            
            # Replace the values in the original dataframes
            objects_df.loc[objects_df['Unnamed: 0'] == slice_number, region] = new_value_objects
            regions_df.loc[regions_df['Unnamed: 0'] == slice_number, region] = new_value_regions
            
    #print(modifications_objects, modifications_regions)

    # Set "Unnamed: 0" as the index for the dataframes
    objects_df.set_index("Unnamed: 0", inplace=True)
    regions_df.set_index("Unnamed: 0", inplace=True)
    colliculi_objects_df.set_index("Unnamed: 0", inplace=True)
    colliculi_regions_df.set_index("Unnamed: 0", inplace=True)

    # Extract the modified cells for highlighting
    modified_cells_objects = {k[0]: [k[1]] for k in modifications_objects.keys()}
    modified_cells_regions = {k[0]: [k[1]] for k in modifications_regions.keys()}

    # Save dataframes with highlighted modified cells
    with pd.ExcelWriter(adjusted_filename(objects_path), engine='openpyxl') as writer:
        (objects_df.style.apply(highlight_cells_modified, axis=1)
         .to_excel(writer, sheet_name='Objects', index=True))

    with pd.ExcelWriter(adjusted_filename(regions_path), engine='openpyxl') as writer:
        (regions_df.style.apply(highlight_cells_modified, axis=1)
         .to_excel(writer, sheet_name='Regions', index=True))

    print(f"Dataframes saved to '{adjusted_filename(objects_path)}' and '{adjusted_filename(regions_path)}' with modified cells highlighted.")

else:
    # Set "Unnamed: 0" as the index for the dataframes
    objects_df.set_index("Unnamed: 0", inplace=True)
    regions_df.set_index("Unnamed: 0", inplace=True)

    # Save dataframes without styling
    with pd.ExcelWriter(adjusted_filename(objects_path), engine='openpyxl') as writer:
        objects_df.to_excel(writer, sheet_name='Objects', index=True)

    with pd.ExcelWriter(adjusted_filename(regions_path), engine='openpyxl') as writer:
        regions_df.to_excel(writer, sheet_name='Regions', index=True)

    print(f"Dataframes saved to '{adjusted_filename(objects_path)}' and '{adjusted_filename(regions_path)}' with updated slice labels.")
    
"""# Handling damage
non_null_values_damage = qc_extended_df["damage"].dropna()
num_non_null_values_damage = len(non_null_values_damage)
if num_non_null_values_damage > 0:
   """ 
