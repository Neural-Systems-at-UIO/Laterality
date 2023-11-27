import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

def compute_unweighted_concordance(index, df, li_columns, window_size):
    start_idx = max(index - window_size, 0)
    end_idx = min(index + window_size + 1, len(df))
    
    current_signs = np.sign(df.loc[index, li_columns].values)
    concordances = []
    for idx in range(start_idx, end_idx):
        neighbor_signs = np.sign(df.loc[idx, li_columns].values)
        concordance = np.mean(current_signs == neighbor_signs)
        concordances.append(concordance)
    
    return np.mean(concordances)

def generate_plot(input_file_path, output_folder_path):
    dataset = pd.read_excel(input_file_path)
    labels_to_exclude = ["4th ventricle"]
    dataset = dataset[~dataset['Base Label'].isin(labels_to_exclude)]
    li_columns = [f"{i:02}_LI" for i in range(3, 7)]
    dataset['avg_absolute_LI'] = dataset[li_columns].abs().sum(axis=1) / 4
    ordered_dataset = dataset.sort_values(by='avg_absolute_LI', ascending=False).reset_index(drop=True)
    window = 5
    ordered_dataset[f'unweighted_concordance_w{window}'] = [compute_unweighted_concordance(idx, ordered_dataset, li_columns, window) for idx in ordered_dataset.index]

    fig, ax = plt.subplots(figsize=(20, 10))
    colors = plt.cm.RdBu_r(ordered_dataset[f'unweighted_concordance_w{window}'])
    bars = ax.bar(ordered_dataset['Base Label'], ordered_dataset['avg_absolute_LI'], color=colors)
    ax.axhline(0, color='gray', linestyle='--', label='Zero Difference')
    ax.set_ylabel('Average Absolute Laterality Index (LI)')
    ax.set_xlabel('Subregion')
    ax.legend()
    ax.set_title(f'Unweighted Concordance with Window Size = {window}')
    ax.set_xticklabels(ordered_dataset['Base Label'], rotation=45, ha='right', fontsize=8)
    cbar = plt.colorbar(plt.cm.ScalarMappable(cmap=plt.cm.RdBu_r), ax=ax, orientation='vertical')
    cbar.set_label(f'Unweighted Concordance with Window Size = {window}')
    plt.tight_layout(pad=3.0)
    plt.savefig(f"{output_folder_path}/UnweightedConcordanceAALI.png")

if __name__ == "__main__":
    input_file_path = input("Please enter the input file path: ")
    output_folder_path = input("Please enter the output folder path where the image will be saved: ")
    generate_plot(input_file_path, output_folder_path)
    print("Plot successfully saved!")
