import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import xlsxwriter

def prepare_data(df):
    flattened = df.values.flatten()
    return flattened[~np.isnan(flattened)]

def plot_distribution(df_without_totals, df_totals, rat_number, dir_path):

    data_without_totals = prepare_data(df_without_totals)
    data_totals = prepare_data(df_totals)
    
    # Compute mean and standard deviation on the flattened data without totals
    mean_without_totals = np.mean(data_without_totals)
    std_without_totals = np.std(data_without_totals)

    # Compute mean and standard deviation on the flattened totals data
    mean_totals = np.mean(data_totals)
    std_totals = np.std(data_totals)
  
    # Calculate outlier thresholds
    upper_threshold_without_totals = mean_without_totals + 2 * std_without_totals
    lower_threshold_without_totals = mean_without_totals - 2 * std_without_totals

    upper_threshold_totals = mean_totals + 2 * std_totals
    lower_threshold_totals = mean_totals - 2 * std_totals

    # Print the means and SDs
    print("Statistics for df_without_totals:")
    print("Mean:", mean_without_totals)
    print("SD:", std_without_totals)

    print("\nStatistics for df_totals:")
    print("Mean:", mean_totals)
    print("LSD:", std_totals)

    # Print the thresholds
    print("\nOutlier Thresholds for df_without_totals:")
    print("Upper Threshold:", upper_threshold_without_totals)
    print("Lower Threshold:", lower_threshold_without_totals)

    print("\nOutlier Thresholds for df_totals:")
    print("Upper Threshold:", upper_threshold_totals)
    print("Lower Threshold:", lower_threshold_totals)

    # Create subplots with 2 rows
    fig, axs = plt.subplots(2, 1, figsize=(10, 8))

    # Plot for data without totals
    _, bins, _ = axs[0].hist(data_without_totals, bins=100, color='grey', edgecolor='black', alpha=0.5)
    lower_bound_2sd = mean_without_totals - 2*std_without_totals
    upper_bound_2sd = mean_without_totals + 2*std_without_totals
    axs[0].axvline(x=lower_bound_2sd, color='red', linestyle='--', linewidth=2)
    axs[0].axvline(x=upper_bound_2sd, color='red', linestyle='--', linewidth=2)
    axs[0].fill_betweenx(y=[0, axs[0].get_ylim()[1]], x1=lower_bound_2sd, x2=upper_bound_2sd, color='red', alpha=0.2)
    axs[0].set_title(f'{rat_number} Object Size Distribution Without Totals (2 SDs highlighted)')
    axs[0].set_xlabel('Object Size')
    axs[0].set_ylabel('Frequency')
    axs[0].grid(True)

    # Plot for totals data
    axs[1].hist(data_totals, bins=bins, color='grey', edgecolor='black', alpha=0.5)
    lower_bound_2sd = mean_totals - 2*std_totals
    upper_bound_2sd = mean_totals + 2*std_totals
    axs[1].axvline(x=lower_bound_2sd, color='red', linestyle='--', linewidth=2)
    axs[1].axvline(x=upper_bound_2sd, color='red', linestyle='--', linewidth=2)
    axs[1].fill_betweenx(y=[0, axs[1].get_ylim()[1]], x1=lower_bound_2sd, x2=upper_bound_2sd, color='red', alpha=0.2)
    axs[1].set_title(f'{rat_number} Object Size Distribution Totals (2 SDs highlighted)')
    axs[1].set_xlabel('Object Size')
    axs[1].set_ylabel('Frequency')
    axs[1].grid(True)

    # Adjust layout
    plt.tight_layout()

    # Save the plot
    plt.savefig(os.path.join(dir_path, f'{rat_number}_Object_Size_Distribution_2_SDs_highlighted.png'))

    # Show the plot
    plt.show()

    # Return the outlier data
    return upper_threshold_without_totals, lower_threshold_without_totals, upper_threshold_totals, lower_threshold_totals

def apply_outlier_formatting(df, upper_threshold, lower_threshold, upper_threshold_last_row, lower_threshold_last_row, file_path):
    # Setup an ExcelWriter object
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1')

        # Access the workbook and worksheet objects from the ExcelWriter
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        # Define formats
        outlier_format = workbook.add_format({'bg_color': 'orange'})
        last_row_format = workbook.add_format({'bg_color': 'pink'})

        # Get the dimensions of the dataframe
        (max_row, max_col) = df.shape

        # Initialize a list to store the positions of the outliers
        outlier_indices = []

        # Iterate over each cell in the dataframe and apply formats
        for col in range(max_col):
            for row in range(max_row-1):  # Avoiding the last row
                cell_value = df.iloc[row, col]
                if cell_value > upper_threshold or cell_value < lower_threshold:
                    worksheet.write(row + 1, col + 1, cell_value, outlier_format)
                    outlier_indices.append((row, col))

        # Format the last row separately
        for col in range(max_col):
            cell_value = df.iloc[-1, col]
            if cell_value > upper_threshold_last_row or cell_value < lower_threshold_last_row:
                worksheet.write(max_row, col + 1, cell_value, last_row_format)
                outlier_indices.append((max_row-1, col))  # Save the position of the last row separately

        # Return the list of outlier indices
        return outlier_indices
    
def apply_same_formatting_to_counts(outlier_indices, df, file_path):
    # Setup an ExcelWriter object
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1')

        # Access the workbook and worksheet objects from the ExcelWriter
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        # Define formats
        outlier_format = workbook.add_format({'bg_color': 'orange'})
        last_row_format = workbook.add_format({'bg_color': 'pink'})

        # Get the dimensions of the dataframe
        (max_row, max_col) = df.shape

        # Apply formatting based on the outlier indices from sizes_df
        for (row, col) in outlier_indices:
            if row < max_row-1:
                cell_value = df.iloc[row, col]
                worksheet.write(row + 1, col + 1, cell_value, outlier_format)
            else:
                cell_value = df.iloc[row, col]
                worksheet.write(row + 1, col + 1, cell_value, last_row_format)

        # The writer object is automatically closed and the file is saved when exiting the 'with' block
        
def main():
    # Get user input
    rat_number = input("Enter the rat number: ")
    sizes_filepath = input("Enter the path for the object sizes file: ")
    counts_filepath = input("Enter the path for the object counts file: ")
    dir_path = os.path.dirname(sizes_filepath)

    # Load the data into DataFrames and set first column as index
    sizes_df = pd.read_excel(sizes_filepath)
    counts_df = pd.read_excel(counts_filepath)

    sizes_df_no_index = sizes_df.drop(columns=['Unnamed: 0'])
    counts_df_no_index = counts_df.drop(columns=['Unnamed: 0'])

    sizes_df.set_index("Unnamed: 0", inplace=True)
    counts_df.set_index("Unnamed: 0", inplace=True)

    # Splitting the sizes DataFrame into two DataFrames
    sizes_df_without_totals = sizes_df_no_index.iloc[:-1]  # All rows except the last
    sizes_df_totals = sizes_df_no_index.iloc[-1:]  # Only the last row (totals)

    # Splitting the counts DataFrame into two DataFrames
    counts_df_without_totals = counts_df_no_index.iloc[:-1]  # All rows except the last
    counts_df_totals = counts_df_no_index.iloc[-1:]  # Only the last row (totals)

    # Capture the outlier indices from the plot_distribution function
    upper_threshold_without_totals, lower_threshold_without_totals, upper_threshold_totals, lower_threshold_totals = plot_distribution(sizes_df_without_totals, sizes_df_totals, rat_number, dir_path)

    # After capturing the outlier thresholds, call the updated function
    sizes_updated_path = os.path.join(dir_path, f'{rat_number}_sizes_with_outliers_highlighted.xlsx')
    outlier_indices = apply_outlier_formatting(sizes_df, upper_threshold_without_totals, lower_threshold_without_totals, upper_threshold_totals, lower_threshold_totals, sizes_updated_path)
    
    # Now apply the same formatting to counts_df
    counts_updated_path = os.path.join(dir_path, f'{rat_number}_counts_with_outliers_highlighted.xlsx')
    apply_same_formatting_to_counts(outlier_indices, counts_df, counts_updated_path)
    
# Call the main function
if __name__ == "__main__":
    main()
