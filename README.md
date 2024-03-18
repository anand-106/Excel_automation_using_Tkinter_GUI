
# Data Entry Application with Tkinter and Openpyxl

This is a simple data entry application developed using Python's Tkinter library for GUI and Openpyxl for Excel file manipulation. The application allows users to enter data, search for existing data, and update an Excel spreadsheet dynamically.


## Features

- User Entry: Users can enter their name and specify the amount for different months.
- User Search: Users can search for existing names and update their corresponding amounts for different months.
- Excel Integration: The entered data is stored in an Excel spreadsheet (Book1.xlsx), where each row represents a user and each column represents a month.
- Dynamic Updates: The application dynamically updates the Excel sheet when new data is entered or existing data is updated.

## Requirements

- Python 3.x
- Tkinter
- Openpyxl
## Installation and Usage

1. Clone the repository to your local machine:

```bash
  git clone https://github.com/your_username/your_repository.git
```
2. Install the required dependencies:

```bash
  pip install openpyxl
```
3. Run the application:

```bash
  python data_entry_app.py
```
## Usage Instructions
- Enter a name and specify the amount for different months in the "User Entry" section.
- Use the "Confirm" button to add a new entry.
- Search for existing names in the "User Search" section and update their corresponding amounts for different months.
- Use the "Confirm" button in the "User Search" section to update the amount.

## Acknowledgements

 - Inspired by [Tkinter documentation](https://docs.python.org/3/library/tkinter.html)
 - Built using [Openpyxl documentation](https://openpyxl.readthedocs.io/en/stable/)

