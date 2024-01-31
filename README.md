# Savi Pricelist Updater

## Project Overview
The Excel Price Updater is a Python-based tool that allows users to update product prices in one Excel file based on the prices listed in another. It utilizes a graphical user interface (GUI) created with Tkinter, making it user-friendly and easily accessible for non-technical users. The key feature of this application is its ability to handle Excel files where the column names for product codes and prices may vary.

## Features
- Browse and select two Excel files for comparison.
- Manually input the column names for product codes and prices.
- Choose the output folder for saving the updated Excel file.
- Display the count and details of updated prices in the GUI.

## Installation
### Prerequisites
- Python 3.x
- Pandas Library
- Openpyxl Library
- Tkinter (usually included in standard Python installations)

### Setup
1. Ensure Python 3.x is installed on your system. If not, download and install it from [Python's official website](https://www.python.org/downloads/).
2. Install the required Python libraries. You can do this by running the following commands in your Python environment:
   ```bash
   pip install pandas openpyxl
   ```
3. Download the project files to your local machine.

## Usage
1. Run the script:
   ```bash
   python excel_price_updater.py
   ```
2. In the GUI:
   - Click the "Browse" buttons to select the two Excel files for price comparison.
   - Enter the column names for product codes and prices as they appear in your Excel files.
   - Click the "Browse" button in the Output Folder section to select where the updated file will be saved.
   - Click "Process Files" to start the updating process.
   - Review the results in the text area of the GUI.

## Note
- Ensure that the Excel files are correctly formatted and accessible.
- Column names are case-sensitive. Enter them exactly as they appear in the Excel files.


