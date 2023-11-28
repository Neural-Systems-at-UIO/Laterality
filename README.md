# Laterality
Python code for performing laterality analysis of rat and mouse brain data

### Python Scripts for Data Processing (Data_Cleaning_Scripts)
1. `SliceCellCountMerger.py` Merges cell counts into a single table, with slices as rows and regions as columns. If a slice is missing, an empty line is created. The
output tables are called “[Rat_number]_[col]_objects.xlsx”.
2. `SliceRegionPixelMerger.py`: Merges region pixels into a single table, with slices as rows and regions as columns. If a slice is missing, an empty line is created.
The output tables are called “[Rat_number]_[col]_regions.xlsx”.
3. `QC.py**: Extends the regions contained in the qc file to the full list of their subregions, as per “brain_regions_dict.json”, resulting in a “[Rat_number]_qc_extended.xlsx” file.
4. `HandlingPartsAdjustments.py`:  Allows for retrieval of cells requiring “parts adjustment” from the “[Rat_number]_qc_extended.xlsx” file and replaces them with the
correct values from the colliculi cell counts and colliculi region pixels tables. It also highlights them in yellow (“00FFFF00”).
5. `HandlingDamage.py`: Allows for automated averaging of damaged data across no more than two adjacent slices for a given region.
If the slice is the first or the last, the data of the closest undamaged slice is copied. Whenever the damaged neighboring slices are more than 2, their data is excluded as missing.
The corrected/excluded data is highlighted in the final cell counts and region pixels tables in dark green (“00008000”). These two last codes are applied sequentially to all the rats.
The files “[Rat_number]_objects_adjustments_damage.xlsx” and “[Rat_number]_regions_adjustments_damage.xlsx” contain the final corrected data for the rat.
6. `ObjectsSizingAndCounting.py`: Populates a table for each rat and each colliculi data containing the available recorded average sizes of neurons in a given slice, per region. The
neuronal cell sizes are in pixels. They are absent (nan) when no neurons are observed in a given area of a given slice. The output file “[Rat_number]_objects_sizes.xlsx” contains slices
as rows and regions as columns, and preserves the yellow/green highlights from the “[Rat_number]_objects_adjustments_damage.xlsx” table for further processing. The code also creates a table with the object counts for each region per slice, named “[Rat_number]_objects_counts.xlsx”, highlighted as well.
7. `HandlingSizingAndCountingDamage.py`: Reports the correct sizes and counts of colliculi and to average data between neighbouring slices where no more than two adjacent
slices are damaged. In this case, similarly to what happens when correcting the top and bottom slice, where we take the same neighbouring value rather than an average, also when a damaged region object size is damaged, we report the adjacent non damaged one, if the other is absent. This is because the neuronal cell count of the other non damaged area is 0 and the average neuronal cell count for the damaged area is different from zero, and so must be the expected object size. Instead, if both the neighbouring undamaged cells contain nan for object
dimension (0 object counts), then also the damaged area has missing object size (and 0 count). The output tables are called “[Rat_number]_objects_sizes_corrected.xlsx” and
“[Rat_number]_objects_counts_corrected.xlsx”. This last table is actually equal to the original “[Rat_number]_objects.xlsx”.
8. `SizesAndCountsTotaling.py`: Generates the “[Rat_number]_objects_sizes_corrected_with_totals.xlsx” and “[Rat_number]_objects_counts_corrected_with_totals.xlsx”, an updated version of the previous tables with the added totals for regions (weighted average sizes and sum of counts).

### Additional Scripts for Analysis (Data_Cleaning_Scripts)
9. `OutliersIdentification.py`: Identifies outliers in neuronal cell sizes. Saves the distribution of neuronal cell sizes and highlights in orange (FFFF6600) and pink (FFFF00FF) the outliers for the full data and average regional data, respectively. 
10. `PlottingOutliers.py`: For analysis and interpretation of anomalous data based on cell counts. Plots 4 graphs to analyze and interpret anomalous data based
on cell counts, for full data and average regional data, respectively.
11. `JointOutliersPlotting.py`: Allows for a comprehensive assessment of size outliers with respect to object counts across all four rats.
12. `EstimatingDensity.py`: Computes the densities in cell counts per mm3, both for each sample, and for each region. The output files are named “[Rat_number]_densities.xlsx”.
13. `DensitiesAndCounts.py`: Merges density and object count data across rats in the same file “All_densities.xlsx”.
14. `FinalDensities.py`: Removes non bilateral regions (i.e. Clear Label) and divides R and L data. It also filters row to remove those below a given threshold (200). The output file is “Processed_All_densities.xlsx”.
15. `LateralityIndex.py`: Evaluates the Laterality Index for each rat and brain region. The output file is “Processed_All_densities_with_LI.xlsx”.

### Visualization Scripts (Plotting_Code_and_Images)
1. `ScatterPlots.py`: Generates scatterplots of right versus left densities for each rat.
2. `BlandAltman.py`: Creates Bland-Altman Plots for Laterality Index analysis.
3. `LateralityIndexRanked.py` & `AverageAbsoluteLateralityIndex.py`: Tools for analyzing brain laterality.
4. `AbsLIConcordance.py` & `ConcordanceHeatmap.py`: Scripts for assessing lateralization patterns in brain regions.

# Authors
Luca 

# Acknowledgements
