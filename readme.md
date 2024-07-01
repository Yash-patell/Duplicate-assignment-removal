# Duplicate  Submission Removal Tool

This is a simple GUI application built using customtkinter and pandas to help identify and remove duplicate submissions in an Excel file. The application allows users to select an Excel file, process it to identify duplicate entries, and save the cleaned data to a new file.

## Features
- Simple and intuitive graphical user interface.
- Open and read Excel files.
- Identify duplicate submissions based on specific columns.
- Remove duplicate entries, keeping only the first submission.
- Save the cleaned data to a new Excel file with a unique name if a file with the same name already exists.

## Installation

1. Clone the repository
-git clone https://github.com/Yash-patell/Duplicate-assignment-removal.git

2. Navigate to the project directory

3. Install required dependencies:
-pip install -r requirements.txt

## Usage: 

1. Run the application

2. Click on the "Insert File" button to open a file dialog and select the Excel file you want to process.

3. The selected file path will be displayed in the application.

4. Click the "Start" button to process the file and remove duplicate submissions.

5. If the output file already exists, you will be prompted to either overwrite it or save the cleaned data to a new file with a unique name.

6. A success message will be shown once the file is saved.

7. Click the "Quit" button to close the application.

## Dependencies

1. pandas: Used for data manipulation and analysis.
2. customtkinter: A modern and customizable alternative to the standard 'tkinter'
3. library for creating graphical user interfaces.
4. tkinter: Standard Python interface to the Tk GUI toolkit.

## Contributing
Contributions are welcome! Please open an issue or submit a pull request if you have any improvements /new features or bug fixes.


