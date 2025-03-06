"""
Excel Data Appender Application

This application provides a GUI for users to append data from one Excel file to another.
It allows users to:
1. Select source and target Excel files
2. Map columns between the files
3. Append data from source to target based on the mappings

The application uses PyQt6 for the GUI components and pandas for Excel data handling.
"""

import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QLabel, QFileDialog, QMessageBox
from PyQt6.QtCore import Qt
from mapping_review_dialog import MappingReviewDialog
from data_appender import append_data

#----------------------------------------------------
# Main Application Window
#----------------------------------------------------
class MainWindow(QMainWindow):
    """
    The main application window that provides the user interface for
    selecting files and initiating the data append operation.
    """
    def __init__(self):
        """Initialize the main window and set up the GUI components."""
        super().__init__()
        self.setWindowTitle("Excel Data Appender")
        self.setGeometry(100, 100, 600, 200)

        # Create a central widget to hold all UI components
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        #----------------------------------------------------
        # Source File Selection Section
        #----------------------------------------------------
        source_layout = QHBoxLayout()
        self.source_input = QLineEdit()  # Text field to display selected source path
        source_button = QPushButton("Browse Source")
        # Connect the button click event to the browse_file method
        source_button.clicked.connect(lambda: self.browse_file(self.source_input, "Select Source Excel File"))
        source_layout.addWidget(QLabel("Source File:"))
        source_layout.addWidget(self.source_input)
        source_layout.addWidget(source_button)
        layout.addLayout(source_layout)

        #----------------------------------------------------
        # Target File Selection Section
        #----------------------------------------------------
        target_layout = QHBoxLayout()
        self.target_input = QLineEdit()  # Text field to display selected target path
        target_button = QPushButton("Browse Target")
        # Connect the button click event to the browse_file method
        target_button.clicked.connect(lambda: self.browse_file(self.target_input, "Select Target Excel File"))
        target_layout.addWidget(QLabel("Target File:"))
        target_layout.addWidget(self.target_input)
        target_layout.addWidget(target_button)
        layout.addLayout(target_layout)

        #----------------------------------------------------
        # Action Button Section
        #----------------------------------------------------
        append_button = QPushButton("Append Data")
        # Connect the button click event to the append_data method
        append_button.clicked.connect(self.append_data)
        layout.addWidget(append_button, alignment=Qt.AlignmentFlag.AlignCenter)

    def browse_file(self, input_field, dialog_title):
        """
        Open a file dialog to select an Excel file and update the input field.
        
        Args:
            input_field: The QLineEdit to update with the selected file path
            dialog_title: Title for the file dialog window
        """
        file_name, _ = QFileDialog.getOpenFileName(self, dialog_title, "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            input_field.setText(file_name)

    def append_data(self):
        """
        Process the append operation when the user clicks the 'Append Data' button.
        This method:
        1. Validates that both files are selected
        2. Opens a mapping dialog for column selection
        3. Performs the append operation and displays results
        """
        source_file = self.source_input.text()
        target_file = self.target_input.text()

        # Validate that both files are selected
        if not source_file or not target_file:
            QMessageBox.warning(self, "Missing Files", "Please select both source and target files.")
            return

        # Open the mapping dialog for user to review and adjust column mappings
        mapping_dialog = MappingReviewDialog(source_file, target_file)
        if mapping_dialog.exec():
            manual_mappings = mapping_dialog.get_manual_mappings()
            try:
                # Perform the actual data append operation
                stats = append_data(source_file, target_file, manual_mappings)
                # Display success message with operation statistics
                QMessageBox.information(self, "Operation Successful", 
                    f"Data appended successfully!\n\n"
                    f"Source columns: {stats['source_columns']}\n"
                    f"Target columns: {stats['target_columns']}\n"
                    f"Mapped columns: {stats['mapped_columns']}\n"
                    f"Appended rows: {stats['appended_rows']}"
                )
            except Exception as e:
                # Display error message if operation fails
                QMessageBox.critical(self, "Error", f"An error occurred while appending data: {str(e)}")

#----------------------------------------------------
# Application Startup
#----------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
