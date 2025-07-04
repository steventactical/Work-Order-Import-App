# Work Order Import Tool

A Windows Forms application designed to automate the process of importing work order data into a nesting software. The program allows users to input a work order number and automatically retrieves relevant data from an ERP system and a programming software database, saving time and reducing the chance of user error. It also provides a manual input option when work orders cannot be found.

## Features
- **Work Order Import**: Input a work order number and retrieve relevant data from an ERP system and a programming software database.
- **Part Number Matching**: If part numbers don’t match, the user can manually match them for correct import.
- **Export to CSV**: After importing all work orders, export the list as a CSV file for seamless integration with nesting software.
- **Manual Data Entry**: Allows the user to manually enter work order data if the work order number is not found.

## How to Use
1. **Clone the repository**:
    `git clone https://github.com/yourusername/Work-Order-Import.git`
   
2. **Run the application**:
    - Open the solution in Visual Studio.
    - Build and run the project.
    
3. **Running in Test Environment for Employers**:
    - To run the program in a test environment, **select the demo database** in the settings.
    - The demo database will contain **work orders WO001 - WO005** for test importing.

4. **Features**:
    - Enter a work order number in the provided field.
    - If the work order is found, relevant data will be populated.
    - Manually input data if the work order is not found.
    - Export the data as a CSV file.

## Technologies Used
- C#
- Windows Forms
- SQL Server (ERP system)
- SQLite (Programming software database)
