"""
Data Appender Module

This module provides functionality to append data from a source Excel file
to a target Excel file based on user-defined column mappings.

The main function takes care of:
1. Reading data from both files
2. Applying column mappings to align data structure
3. Appending mapped data to the target file
4. Generating statistics about the operation
"""

import pandas as pd

def append_data(source_file, target_file, manual_mappings):
    """
    Append data from source Excel file to target Excel file using specified column mappings.
    
    Args:
        source_file (str): Path to the source Excel file
        target_file (str): Path to the target Excel file
        manual_mappings (dict): Dictionary of {target_column: source_column} mappings
    
    Returns:
        dict: Statistics about the operation including column and row counts
    
    Raises:
        Various exceptions if file access or data processing fails
    """
    # Step 1: Read both Excel files into pandas DataFrames
    source_df = pd.read_excel(source_file)
    target_df = pd.read_excel(target_file)

    # Step 2: Create a copy of source data to apply mappings
    mapped_source_df = source_df.copy()
    
    # Apply column mappings by renaming source columns to match target columns
    for target_col, source_col in manual_mappings.items():
        if source_col in mapped_source_df.columns:
            mapped_source_df = mapped_source_df.rename(columns={source_col: target_col})

    # Step 3: Find columns that exist in both dataframes after mapping
    common_columns = list(set(mapped_source_df.columns) & set(target_df.columns))

    # Step 4: Clean data - Fill NaN values with empty strings to avoid errors
    for col in common_columns:
        mapped_source_df[col] = mapped_source_df[col].fillna('') 

    # Step 5: Append data - only using columns that exist in the target file
    appended_df = pd.concat([target_df, mapped_source_df[common_columns]], ignore_index=True)

    # Step 6: Save the combined data back to the target file
    appended_df.to_excel(target_file, index=False)

    # Step 7: Return statistics about the operation
    return {
        "source_columns": len(source_df.columns),
        "target_columns": len(target_df.columns),
        "mapped_columns": len(common_columns),
        "appended_rows": len(mapped_source_df)
    }