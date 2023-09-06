import pandas as pd
import matplotlib.pyplot as plt

def plot_li_difference():
    # Prompt user for input and output paths
    input_path = input("Please enter the path to the input Excel file: ")
    output_folder = input("Please enter the path to the output folder: ")
    
    # Load the data from the Excel file
    df = pd.read_excel(input_path)

    # Set up the figure and axes for the plots
    fig, axs = plt.subplots(2, 2, figsize=(16, 14))
    fig.suptitle('Difference Plots of Laterality Index (LI) with Means and 1.96 SDs', fontsize=16)
    
    rats = ["Rat03", "Rat04", "Rat05", "Rat06"]
    li_columns = [f"LI_{i:02}" for i in range(3, 7)]

    # Plot LI vs. average density with means and 1.96 SDs for each rat
    for i, ax in enumerate(axs.ravel()):
        li_column = li_columns[i]
        rat = rats[i]
        
        # Calculate the average density
        averages = (df[f"{rat}_p_L"] + df[f"{rat}_p_R"]) / 2
        
        # Scatter plot
        ax.scatter(averages, df[li_column], label='Data', alpha=0.6, color='green')
        
        # Highlight zero difference with a horizontal line
        ax.axhline(0, color='gray', linestyle='--', label='Zero Difference')
        
        # Calculate mean and 1.96 SDs for LI
        mean_li = df[li_column].mean()
        sd_li = df[li_column].std()
        
        # Add mean and 1.96 SDs to the plot
        ax.axhline(mean_li, color='red', linestyle='-', label='Mean LI')
        ax.axhline(mean_li + 1.96 * sd_li, color='blue', linestyle='--', label='Mean LI + 1.96 SDs')
        ax.axhline(mean_li - 1.96 * sd_li, color='blue', linestyle='--', label='Mean LI - 1.96 SDs')
        
        # Set title, labels, and legend
        ax.set_title(rat)
        ax.set_xlabel('Average Density (L + R) / 2')
        ax.set_ylabel('Laterality Index (LI)')
        ax.legend()

    plt.tight_layout(pad=3.0)
    
    # Save the figure to the specified output folder
    output_path = f"{output_folder}/LI_difference_plot.png"
    plt.savefig(output_path)
    print(f"Plot saved to: {output_path}")

if __name__ == "__main__":
    plot_li_difference()
