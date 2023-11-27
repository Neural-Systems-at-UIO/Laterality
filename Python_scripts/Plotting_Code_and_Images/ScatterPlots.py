import pandas as pd
import matplotlib.pyplot as plt
import os

def create_scatter_plot(input_file, output_folder):
    # Load the dataset
    df = pd.read_excel(input_file)
    
    # List of rat numbers
    rats = ['03', '04', '05', '06']
    
    # Filtering criteria to compute the regression without outliers or non-relevant regions
    labels_to_exclude = ["Reticular (pre)thalamic nucleus, auditory segment", 
                         "Reticular (pre)thalamic nucleus, unspecified",
                         "external medullary lamina, auditory radiation",
                         "4th ventricle"]
    filtered_df = dataset = df[~df['Base Label'].isin(labels_to_exclude)]
    
    # Set up the figure for the scatter plots
    fig, axs = plt.subplots(2, 2, figsize=(15, 12))
    fig.suptitle('Scatter Plots of Densities for Subregions (With Zero-Constrained Regression): Left vs. Right', fontsize=16)

    # Plot scatter plots and regression lines constrained to pass through the origin for each rat's densities
    for i, ax in enumerate(axs.ravel()):
        rat = rats[i]

        # Extract data for the rat
        x_data = filtered_df[f"{rat}_densities R"]
        y_data = filtered_df[f"{rat}_densities L"]

        # Scatter plot
        ax.scatter(x_data, y_data, label='Data', alpha=0.6, marker='x')

        # Bisector (dotted line for symmetry)
        ax.plot([0, max(x_data.max(), y_data.max())], 
                [0, max(x_data.max(), y_data.max())], 
                '--', color='gray', label='Bisector')

        # Zero-constrained linear regression
        slope = (x_data * y_data).sum() / (x_data**2).sum()
        ax.plot(x_data, slope * x_data, color='red', label=f"Zero-Constrained Regression")

        # Set title, labels, and legend
        ax.set_title("Rat " + rat)
        ax.set_xlabel('Density Right')
        ax.set_ylabel('Density Left')
        ax.legend()

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    plt.savefig(os.path.join(output_folder, 'ScatterPlots.png'))

if __name__ == "__main__":
    # Get user input for file paths
    input_file = input("Enter the path to the input Excel file: ")
    output_folder = input("Enter the path to the output folder where the image should be saved: ")
    
    # Create and save the scatter plot
    create_scatter_plot(input_file, output_folder)
    print("Scatter plot saved successfully!")
