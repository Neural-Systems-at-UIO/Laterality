import os
import pandas as pd

def get_file_label(file_path, is_density=True):
    # Extract the base file name without the full path
    base_name = os.path.basename(file_path)
    # Remove the file extension
    base_name = base_name.split('.')[0]  # This line removes the file extension
    # If it's a counts file, truncate after "counts"
    if not is_density:
        base_name = base_name.split("counts")[0] + "counts"
    # Return the modified base name as the label
    return base_name


def load_file_and_get_last_row(file_path, is_density=True):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path, index_col='Unnamed: 0')
    # Get the last row of the DataFrame
    last_row = df.iloc[-1]
    # Get the file label
    file_label = get_file_label(file_path, is_density)
    return last_row, file_label

def main():
    # Lists to store the last rows and labels
    densities_last_rows = []
    densities_labels = []
    counts_last_rows = []
    counts_labels = []
    
    # Prompt for density files and load the last row
    for i in range(4):
        file_path = input(f"Enter the path for density file {i+1}: ")
        last_row, label = load_file_and_get_last_row(file_path, is_density=True)
        densities_last_rows.append(last_row)
        densities_labels.append(label)
    
    # Prompt for counts files and load the last row
    for i in range(4):
        file_path = input(f"Enter the path for counts file {i+1}: ")
        last_row, label = load_file_and_get_last_row(file_path, is_density=False)
        counts_last_rows.append(last_row)
        counts_labels.append(label)

    # Combine all last rows into a DataFrame
    combined_data = pd.concat([pd.Series(row) for row in densities_last_rows + counts_last_rows], axis=1)
    
    # Use the labels as column names
    combined_data.columns = densities_labels + counts_labels
    
    # Extract the common header from the first density file (assuming it's common across all files)
    header_file_path = input("Enter the path for the header file: ")
    header_df = pd.read_excel(header_file_path, index_col='Unnamed: 0')
    header = header_df.columns.tolist()
    
    # Insert the header as the first column in the combined DataFrame
    combined_data.insert(0, 'Label', header)
    
    save_path = input('Enter the directory path where you want to save the All_densities.xlsx file: ')

    # Save the combined DataFrame to an Excel file
    combined_data.to_excel(f'{save_path}/All_densities.xlsx', index=False)

    print("The combined Excel file has been saved as 'All_densities.xlsx'")
    
if __name__ == "__main__":
    main()
