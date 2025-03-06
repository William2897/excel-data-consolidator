"""
Column Mapping Dialog Module

This module provides a dialog interface that allows users to:
1. Review automatically detected column mappings between source and target Excel files
2. Manually adjust mappings with the help of fuzzy matching suggestions
3. Confirm or cancel the mapping operation

The dialog uses fuzzy matching to suggest potential column matches when exact matches aren't found.
"""

from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, 
                            QTreeWidget, QTreeWidgetItem, QHeaderView,
                             QComboBox)
import pandas as pd
from fuzzywuzzy import process  # Library for fuzzy string matching

#----------------------------------------------------
# Mapping Review Dialog
#----------------------------------------------------
class MappingReviewDialog(QDialog):
    """
    A dialog window that shows column mappings between source and target Excel files
    and allows the user to adjust them before proceeding with data append operation.
    """
    def __init__(self, source_file, target_file):
        """
        Initialize the mapping review dialog.
        
        Args:
            source_file: Path to the source Excel file
            target_file: Path to the target Excel file
        """
        super().__init__()
        self.setWindowTitle("Review Column Mappings")
        self.setGeometry(200, 200, 800, 600)

        # Store file paths and initialize mapping dictionary
        self.source_file = source_file
        self.target_file = target_file
        self.manual_mappings = {}  # Will store {target_column: source_column} mappings

        # Create the main layout
        layout = QVBoxLayout(self)

        #----------------------------------------------------
        # Column Mapping Tree Section
        #----------------------------------------------------
        # Create a tree widget to display the column mappings
        self.mapping_tree = QTreeWidget()
        self.mapping_tree.setHeaderLabels(["Target Column", "Source Column", "Suggestions"])
        self.mapping_tree.setColumnCount(3)
        # Make columns resize to fill available space
        self.mapping_tree.header().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.mapping_tree)

        #----------------------------------------------------
        # Button Section
        #----------------------------------------------------
        button_layout = QHBoxLayout()
        
        # Proceed button - confirms mappings and closes dialog with accept result
        proceed_button = QPushButton("Proceed")
        proceed_button.clicked.connect(self.accept)
        button_layout.addWidget(proceed_button)
        
        # Cancel button - closes dialog with reject result
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)
        
        layout.addLayout(button_layout)

        # Load and display the initial column mappings
        self.load_mappings()

    def load_mappings(self):
        """
        Load column headers from both Excel files and create initial mappings.
        Creates a tree widget entry for each target column with:
        - Automatic mapping if column names match exactly
        - Suggested mappings based on fuzzy string matching for non-matching columns
        """
        # Read only headers (no data) from both files for efficiency
        source_df = pd.read_excel(self.source_file, nrows=0)
        target_df = pd.read_excel(self.target_file, nrows=0)

        source_columns = source_df.columns
        target_columns = target_df.columns

        # For each target column, try to find a matching source column
        for target_col in target_columns:
            # Create a tree item for this target column
            item = QTreeWidgetItem(self.mapping_tree)
            item.setText(0, target_col)  # Set target column name
            
            # If exact match found, use it as the default mapping
            if target_col in source_columns:
                item.setText(1, target_col)
                self.manual_mappings[target_col] = target_col
            else:
                # No exact match, mark as not mapped and provide suggestions
                item.setText(1, "Not Mapped")
                
                # Use fuzzy matching to find potential matches (min 60% similarity)
                matches = process.extract(target_col, source_columns, limit=5)
                suggestions = [match[0] for match in matches if match[1] >= 60]
                
                # Create dropdown with suggestions
                combo = QComboBox()
                combo.addItem("Select mapping...")
                combo.addItems(suggestions)
                # Connect change event to update mapping when user selects an option
                combo.currentIndexChanged.connect(lambda index, item=item, combo=combo: 
                                                 self.on_mapping_selected(item, combo))
                
                # Add the dropdown to the tree widget in the suggestions column
                self.mapping_tree.setItemWidget(item, 2, combo)

    def on_mapping_selected(self, item, combo):
        """
        Update mapping when user selects an option from the dropdown.
        
        Args:
            item: The QTreeWidgetItem representing the row
            combo: The QComboBox with the selection
        """
        # Only update if user actually selects a mapping (not the placeholder)
        if combo.currentIndex() > 0:
            target_col = item.text(0)
            source_col = combo.currentText()
            
            # Update display and internal mapping dictionary
            item.setText(1, source_col)
            self.manual_mappings[target_col] = source_col

    def get_manual_mappings(self):
        """
        Return the final column mappings after user review.
        
        Returns:
            dict: Dictionary of {target_column: source_column} mappings
        """
        return self.manual_mappings
