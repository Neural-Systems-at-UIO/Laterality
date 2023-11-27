import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

def compute_unweighted_concordance(index, df, li_columns, window_size):
    # Finding the actual positions of the index
    actual_idx_pos = df.index.get_loc(index)
    
    # Get the range for the window based on actual positions
    start_pos = max(actual_idx_pos - window_size, 0)
    end_pos = min(actual_idx_pos + window_size + 1, len(df))
    
    # Get the actual index values for the window range
    window_indices = df.index[start_pos:end_pos]
    
    current_signs = np.sign(df.loc[index, li_columns].values)
    concordances = []
    for idx in window_indices:
        neighbor_signs = np.sign(df.loc[idx, li_columns].values)
        concordance = np.mean(current_signs == neighbor_signs)
        concordances.append(concordance)
    return np.mean(concordances)

def get_clusters(dataframe, li_columns):
    # Calculate unweighted concordance for each region
    window = 5
    # Calculate the threshold for high concordance
    high_concordance_threshold = dataframe[f'unweighted_concordance_w{window}'].quantile(0.75)

    # Identify clusters of regions with high concordance
    clusters = []
    current_cluster = []
    for idx, row in dataframe.iterrows():
        if row[f'unweighted_concordance_w{window}'] >= high_concordance_threshold:
            current_cluster.append(row['Base Label'])
        else:
            if current_cluster:
                clusters.append(current_cluster)
                current_cluster = []

    # Add the last cluster if it exists
    if current_cluster:
        clusters.append(current_cluster)
    return clusters

def main():
    # Ask the user for the input file path
    input_path = input("Please enter the input file path: ")

    # Read the dataset and order it by average absolute LIs
    dataset = pd.read_excel(input_path)
    labels_to_exclude = ["4th ventricle"]
    dataset = dataset[~dataset['Base Label'].isin(labels_to_exclude)]
    dataset['avg_absolute_LI'] = dataset[['03_LI', '04_LI', '05_LI', '06_LI']].abs().mean(axis=1)
    ordered_dataset = dataset.sort_values('avg_absolute_LI', ascending=False)
    li_columns = ['03_LI', '04_LI', '05_LI', '06_LI']
    
    # Calculate unweighted concordance for each region and add it to dataframe
    window_size = 5
    ordered_dataset[f'unweighted_concordance_w{window_size}'] = ordered_dataset.apply(
        lambda row: compute_unweighted_concordance(row.name, ordered_dataset, li_columns, window_size), axis=1)

    # Get the clusters
    clusters = get_clusters(ordered_dataset, li_columns)

    # Flatten the clusters to get a list of all unique regions
    all_clustered_regions = [region for cluster in clusters for region in cluster]

    # Create a matrix to store pairwise concordance scores
    num_regions = len(all_clustered_regions)
    concordance_matrix = np.zeros((num_regions, num_regions))

    # Calculate pairwise concordance scores for all regions
    for i in range(num_regions):
        for j in range(num_regions):
            region_a = all_clustered_regions[i]
            region_b = all_clustered_regions[j]
            li_a = ordered_dataset.loc[ordered_dataset['Base Label'] == region_a, li_columns].values[0]
            li_b = ordered_dataset.loc[ordered_dataset['Base Label'] == region_b, li_columns].values[0]
            concordance = np.mean(np.sign(li_a) == np.sign(li_b))
            concordance_matrix[i, j] = concordance

    # Order the regions based on their average concordance with other regions
    average_concordance = concordance_matrix.mean(axis=1)
    sorted_indices = np.argsort(-average_concordance)

    # Reorder the concordance matrix and region labels
    sorted_matrix = concordance_matrix[sorted_indices][:, sorted_indices]
    sorted_regions = np.array(all_clustered_regions)[sorted_indices]

    # Plot the reordered concordance matrix
    plt.figure(figsize=(16, 14))
    plt.imshow(sorted_matrix, cmap="RdBu_r", aspect='auto')
    plt.colorbar(label='Concordance Score')
    plt.title('Pairwise Concordance Scores Between Regions (Ordered)')
    plt.xticks(np.arange(num_regions), sorted_regions, rotation=45, ha='right', fontsize=10)
    plt.yticks(np.arange(num_regions), sorted_regions, rotation=0, fontsize=10)
    plt.grid(which='both', color='gray', linestyle='--', linewidth=0.5)
    plt.tight_layout()

    # Ask the user for the output folder path
    output_folder = input("Please enter the output folder path: ")
    output_path = f"{output_folder}/ConcordanceHeatmap.png"
    plt.savefig(output_path)
    print(f"Saved the heatmap to {output_path}")

if __name__ == "__main__":
    main()
