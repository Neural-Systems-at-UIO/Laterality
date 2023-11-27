import pandas as pd
import matplotlib.pyplot as plt

def non_overlapping_labels(ax, spacing=0.1):
    prev_label = None
    for label in ax.get_yticklabels():
        if prev_label is not None:
            if abs(label.get_position()[1] - prev_label.get_position()[1]) < spacing:
                label.set_visible(False)
        prev_label = label

def plot_laterality_index(input_file_path, output_folder_path):
    # Read the data
    df = pd.read_excel(input_file_path)
    # Filtering criteria to compute the regression without outliers or non-relevant regions
    labels_to_exclude = ["4th ventricle"]
    df = df[~df['Base Label'].isin(labels_to_exclude)]

    # LI columns
    li_columns = [f"{i:02}_LI" for i in range(3, 7)]

    # Set up the figure and axes for the non-overlapping LI plots
    fig, axs = plt.subplots(2, 2, figsize=(16, 14))
    fig.suptitle('Laterality Index (LI) for Subregions Ordered by LI Magnitude (Non-overlapping Labels)', fontsize=16)

    # Plot LI for each rat, with subregions ordered by LI magnitude and non-overlapping labels
    for i, ax in enumerate(axs.ravel()):
        li_column = li_columns[i]
        rat = f"Rat {i+3:02}"
        
        # Order the subregions by absolute LI magnitude
        ordered_df = df.sort_values(by=li_column, key=abs, ascending=False)
        
        # Plot
        ax.barh(ordered_df['Base Label'], ordered_df[li_column], color='green', alpha=0.6)
        
        # Highlight zero difference with a vertical line
        ax.axvline(0, color='gray', linestyle='--', label='Zero Difference')
        
        # Set title, labels, and legend
        ax.set_title(rat)
        ax.set_xlabel('Laterality Index (LI)')
        ax.set_ylabel('Subregion')
        ax.legend()
        ax.invert_yaxis()
        
        # Adjust y-axis labels for non-overlapping
        ax.set_yticklabels(ordered_df['Base Label'], fontsize=5)
        non_overlapping_labels(ax)

    plt.tight_layout(pad=3.0)
    plt.savefig(f"{output_folder_path}/LIRanked.png")
    print("LI Ranked Plot saved successfully!")

if __name__ == "__main__":
    input_file_path = input("Enter the path to the input Excel file: ")
    output_folder_path = input("Enter the path to the output folder where the image will be saved: ")
    plot_laterality_index(input_file_path, output_folder_path)
