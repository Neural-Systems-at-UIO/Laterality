"""
This code merges all the cell counts for a given rat in a table
whose rows are slices, and columns are regions.
"""

import os
import pandas as pd

def gather_data_from_folder(folder_path):
    """
    Gathers cell count data from all CSV files in the RefAtlasRegions folder and returns a DataFrame 
    with slice numbers as rows and regions as columns.
    """
    # List all files that match the pattern
    files = [f for f in os.listdir(folder_path) if "RefAtlasRegions__s" in f and f.endswith(".csv")]

    # Sort the files based on the slice number
    files.sort(key=lambda x: int(x.split('__s')[1].split('.csv')[0]))

    # Initialize an empty list to store DataFrames
    df_list = []
    
    # Used to store max slice number for checking missing slices later
    max_slice_num = 0

    for file in files:
        # Extract slice number from the file name
        slice_num = int(file.split('__s')[1].split('.csv')[0])
        max_slice_num = max(max_slice_num, slice_num)
        file_path = os.path.join(folder_path, file)
        try:
            # Read the CSV and set 'Region Name' as the index
            df_slice = pd.read_csv(file_path, delimiter=';')
            df_slice.set_index("Region Name", inplace=True)
            df_slice = df_slice[['Object count']].rename(columns={'Object count': slice_num}).transpose()
            df_list.append(df_slice)
        except Exception as e:
            # Handle possible file read errors
            print(f"Error reading {file_path}. It might be malformed. Please check the file.")
            print("Error message:", str(e))
            with open(file_path, 'r') as f:
                print("First few lines of the file:")
                for _ in range(5):
                    print(f.readline().strip())
            continue

    if not df_list:
        raise ValueError("No dataframes were successfully created from the provided CSVs.")

    final_df = pd.concat(df_list)

    # Add missing slices with empty rows
    for missing_slice in range(1, max_slice_num + 1):
        if missing_slice not in final_df.index:
            final_df.loc[missing_slice] = [None] * len(final_df.columns)

    # Sort DataFrame by slice number after adding missing slices
    final_df.sort_index(inplace=True)

    return len(files), final_df


def main():
    """
    Main function that gathers data from CSV files and outputs the combined data to an XLSX file.
    """
    folder_path = input("Enter the path to the folder containing the CSV files: ")
    num_files, final_df = gather_data_from_folder(folder_path)

    # Save the merged data to an XLSX file
    rat_number = input("Enter the rat number: ")
    colliculi_regions = input("Are the counts of the colliculi regions? (yes/no): ")
    suffix = "_col" if colliculi_regions.lower() == "yes" else ""
    output_filename = f"{rat_number}{suffix}_objects.xlsx"
    output_file = os.path.join(folder_path, output_filename)
    final_df.to_excel(output_file, engine='openpyxl')
    print(f"Processed {num_files} files. Results saved to {output_file}")

if __name__ == "__main__":
    main()
