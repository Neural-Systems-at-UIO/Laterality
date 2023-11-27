import pandas as pd
import matplotlib.pyplot as plt
import os

def generate_li_plot():
    # Ask the user for input and output paths
    input_file_path = input("Please enter the input Excel file path: ")
    output_folder_path = input("Please enter the output folder path where the image should be saved: ")

    # Read the dataset from the provided path
    dataset = pd.read_excel(input_file_path)

    # Filtering
    labels_to_exclude = ["4th ventricle"]
    dataset = dataset[~dataset['Base Label'].isin(labels_to_exclude)]

    # Columns for LI
    li_columns = [f"{i:02}_LI" for i in range(3, 7)]

    # Calculate the sum of the absolute values of LI for each region and divide by 4
    dataset['avg_absolute_LI'] = dataset[li_columns].abs().sum(axis=1) / 4

    # Order the dataset by the average absolute LI
    ordered_dataset = dataset.sort_values(by='avg_absolute_LI', ascending=False)

    # Set up the figure and axes for the average absolute LI plots with horizontal bars
    fig, ax = plt.subplots(figsize=(20, 10))
    fig.suptitle('Average Absolute Laterality Index (LI) for All Subregions (Horizontal Bars)', fontsize=16)

    # Plot with horizontal bars
    ax.bar(ordered_dataset['Base Label'], ordered_dataset['avg_absolute_LI'], color='green', alpha=0.6)

    # Highlight zero difference with a horizontal line
    ax.axhline(0, color='gray', linestyle='--', label='Zero Difference')

    # Set labels and legend
    ax.set_ylabel('Average Absolute Laterality Index (LI)')
    ax.set_xlabel('Subregion')
    ax.legend()

    # Adjust x-axis labels to be oblique
    ax.set_xticklabels(ordered_dataset['Base Label'], rotation=45, ha='right', fontsize=10)

    # Save the plot to the specified output folder
    output_file_path = os.path.join(output_folder_path, "AverageAbsoluteLI.png")
    plt.tight_layout(pad=3.0)
    plt.savefig(output_file_path)
    print(f"Plot saved to: {output_file_path}")

# Execute the function
if __name__ == "__main__":
    generate_li_plot()
