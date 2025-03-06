# Excel Data Consolidator

## Overview
The Excel Data Consolidator is a Python application that allows users to append data from one Excel file to another with intelligent column mapping. The tool offers both a desktop GUI application and a Jupyter notebook interface to accommodate different user preferences.

## Features
- **File Selection**: Easily browse and select source and target Excel files
- **Intelligent Column Mapping**: 
  - Automatic mapping of identically named columns
  - Interactive manual mapping for differently named columns
  - Fuzzy matching suggestions for similar column names
- **Data Append Operation**: Combines data based on your mapping choices
- **Operation Statistics**: View detailed information about the mapping and appending process
- **Multiple Interfaces**: Use either the desktop application or Jupyter notebook

## Components
- [`Data_Consolidator.py`](c:\Users\User\OneDrive\Desktop\MAGNA%20Internship\Consolidator\Data_Consolidator.py): Main desktop GUI application
- [`data_appender.py`](c:\Users\User\OneDrive\Desktop\MAGNA%20Internship\Consolidator\data_appender.py): Core data processing module
- [`mapping_review_dialog.py`](c:\Users\User\OneDrive\Desktop\MAGNA%20Internship\Consolidator\mapping_review_dialog.py): Column mapping interface
- [`Excel Data Consolidator.ipynb`](c:\Users\User\OneDrive\Desktop\MAGNA%20Internship\Consolidator\Excel%20Data%20Consolidator.ipynb): Jupyter notebook implementation

## Requirements
- Python 3.x
- PyQt6 (for desktop application)
- pandas
- fuzzywuzzy
- python-levenshtein (optional, improves fuzzy matching performance)
- Jupyter, ipywidgets, ipyfilechooser (for notebook interface)

## Installation
```bash
# Main dependencies
pip install pandas PyQt6 fuzzywuzzy

# Optional - for better fuzzy matching performance
pip install python-levenshtein

# For notebook interface
pip install jupyter ipywidgets ipyfilechooser
```

## Usage

### Desktop Application
1. Run the main application:
   ```bash
   python Data_Consolidator.py
   ```
2. Click "Browse Source" to select your source Excel file
3. Click "Browse Target" to select your target Excel file
4. Click "Append Data" to open the mapping dialog
5. Review and adjust column mappings as needed
6. Click "Proceed" to append the data

### Jupyter Notebook
1. Launch Jupyter and open `Excel Data Consolidator.ipynb`
2. Run all cells to load the interface
3. Use the file choosers to select your source and target files
4. Click "Map Columns" to begin the mapping process
5. Adjust mappings as needed and click "Proceed"

## How It Works
1. Source and target Excel files are loaded into pandas DataFrames
2. Column mappings are established (automatically or manually)
3. Source columns are renamed to match target columns based on mappings
4. Data is cleaned to prevent issues with null values
5. Mapped source data is appended to the target data
6. The combined data is saved back to the target file
7. Operation statistics are displayed
