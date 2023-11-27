import pandas as pd
import json
import os

def process_qc_file_with_args(qc_path, dict_path):
    """
    Process the qc file based on provided paths and the brain regions dictionary.
    """
    # Load the qc file and the JSON dictionary
    qc_df = pd.read_excel(qc_path, header=0, index_col=0)
    with open(dict_path, 'r') as file:
        brain_regions_dict = json.load(file)

    # Process the "damage" and "parts adjustment" columns
    for col in ["damage", "parts adjustment"]:
        qc_df[col] = qc_df[col].apply(
            lambda x: '; '.join(['; '.join(brain_regions_dict.get(name.strip(), [name.strip()])) for name in str(x).split(';')]) if pd.notna(x) else x
        )

    # Save the modified qc file
    base_dir = os.path.dirname(qc_path)
    output_filename = os.path.splitext(os.path.basename(qc_path))[0] + "_extended.xlsx"
    output_path = os.path.join(base_dir, output_filename)
    qc_df.to_excel(output_path)

    return output_path

if __name__ == "__main__":
    # Ask the user for paths to the required files
    qc_path = input("Please enter the path to the quality control (qc) file: ")
    dict_path = input("Please enter the path to the JSON dictionary (brain_regions_dict): ")

    output_path = process_qc_file_with_args(qc_path, dict_path)
    print(f"Processing completed. The modified qc file has been saved as {output_path}.")
