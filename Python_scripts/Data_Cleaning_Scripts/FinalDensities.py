import pandas as pd
import os

def process_data(file_path, threshold):
    # Load the data
    data = pd.read_excel(file_path)

    # Step 1: Filter rows based on labels ending with " R" or " L"
    filtered_data = data[data['Label'].str.endswith((' R', ' L'))].copy()
    
    # Step 2: Verify pairs of labels
    filtered_data['Base Label'] = filtered_data['Label'].str[:-2]
    base_labels_set = set(filtered_data['Base Label'])
    unpaired_labels = {label for label in base_labels_set if not (
        (label + " R") in filtered_data['Label'].values and 
        (label + " L") in filtered_data['Label'].values)}
    if unpaired_labels:
        raise ValueError(f"Unpaired labels detected: {unpaired_labels}")

    # Step 3: Split data into 'R' and 'L'
    data_r = filtered_data[filtered_data['Label'].str.endswith(' R')].copy()
    data_l = filtered_data[filtered_data['Label'].str.endswith(' L')].copy()
    data_r['Base Label'] = data_r['Base Label'].str.strip()
    data_l['Base Label'] = data_l['Base Label'].str.strip()
    data_r.columns = [col if col == 'Base Label' else col + ' R' for col in data_r.columns]
    data_l.columns = [col if col == 'Base Label' else col + ' L' for col in data_l.columns]
    merged_data = pd.merge(data_r, data_l, on='Base Label', how='inner')
    merged_data = merged_data.drop(['Label R', 'Label L'], axis=1)

    # Ensure the 'Base Label' is the first column
    cols = ['Base Label'] + [col for col in merged_data.columns if col != 'Base Label']
    merged_data = merged_data[cols]

    # Step 4: Add average count column
    count_columns = [col for col in merged_data.columns if 'counts' in col]
    merged_data['Average Counts'] = merged_data[count_columns].mean(axis=1)

    # Step 5: Filter rows based on average count threshold
    final_data = merged_data[merged_data['Average Counts'] >= threshold]

    return final_data

# Main program to run the steps
if __name__ == "__main__":
    file_path_input = input("Please enter the path to the Excel file: ")
    threshold_input = float(input("Please enter the threshold for average counts: "))
    
    try:
        processed_data = process_data(file_path_input, threshold_input)
        
        # Extract directory and file components from the input path
        file_directory = os.path.dirname(file_path_input)
        file_name = os.path.basename(file_path_input)
        
        # Construct the output file path in the same directory with a modified file name
        output_file_name = "Processed_" + file_name
        output_file_path = os.path.join(file_directory, output_file_name)
        
        processed_data.to_excel(output_file_path, index=False)
        print(f"Processed data saved to {output_file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")
