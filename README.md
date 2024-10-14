# Exercise-4-Read-and-Write-Excel-Data

## Aim:
To automate the process of reading data from an Excel file and writing data into another Excel file or a
different sheet using UiPath.

## Equipment Required:
UiPath Studio (installed on a compatible system).
Microsoft Excel application.
Excel file (e.g., "ExcelFile.xlsx") containing data for reading and writing purposes.
Computer with:
Minimum 4 GB RAM.
Minimum 2.0 GHz CPU.
Windows operating system.
.NET Framework 4.6.1 or later.

## Procedure:
Start a New Process
Open UiPath Studio.
Create a new process by selecting Process under the New Project tab.
Name the project UiPath Read and Write Excel Data and click Create to begin.

## Install Excel Activities Package:
If the Excel activities are not already available, go to Manage Packages (click on the Project panel).
Search for UiPath.Excel.Activities in the Official packages.
Click Install and then Save to include the Excel activities in the project.

## Add Excel Application Scope:
In the Activities panel, search for Excel Application Scope. Drag and drop the Excel Application Scope
activity into the main sequence window.
In the Properties pane, specify the path to the Excel file (e.g., "C:\Path\To\Your\ExcelFile.xlsx").
The Excel Application Scope will open the Excel file and keep it open while the robot interacts with it.

## Read Data from Excel:
Inside the Excel Application Scope, search for Read Range in the Activities panel.
Drag the Read Range activity inside the Excel Application Scope. Set the SheetName property to
"Sheet1" (or the sheet from which data needs to be read).
Leave the Range field blank if you want to read all data from the sheet, or specify a range like "A1:C10"
if you want specific cells.
Create a variable named dtExcelData to store the output data, which will be of type DataTable.

## Modify the Data (Optional):
If you wish to process the data after reading it, you can use activities like For Each Row to loop
through the rows of the DataTable.
Example: Use Assign activities to modify certain cell values, calculate totals, or apply filters.

## Write Data into Excel:
After processing or modifying the data, add a Write Range activity under the Excel Application Scope.
Set the SheetName to a different sheet, such as "Sheet2" (or the same sheet, if preferred).
Set the DataTable property to dtExcelData to write the data back to the file.
Specify the Range (e.g., "A1") to indicate where the data should be written.
Ensure the Add Headers option is checked if the data contains column headers that need to be
written.

## Save and Run the Workflow:
Save the project by pressing CTRL+S.
Click Run from the toolbar to execute the workflow.
UiPath will read data from Sheet1, process it (if specified), and then write the data to Sheet2 (or the
same sheet).

## Example Scenario:
Initial Excel File:
Sheet1 contains customer data: Name, Age, City.
Sheet2 is initially empty.

## UiPath Workflow Actions:
The robot reads the data from Sheet1.
(Optional) The data is modified (e.g., age is incremented by 1).
The modified data is written to Sheet2, including the headers

## Expected Outcome:
Sheet2 will now have the same data as Sheet1 (or the modified version), demonstrating how UiPath
can read and write data efficiently.

## UiPath WorkFlow:

![image](https://github.com/user-attachments/assets/768588d1-f69b-481a-aa8c-c279883eb2d0)

![image](https://github.com/user-attachments/assets/d920a8b1-57b0-4e31-a3be-fd67378f5e87)

![image](https://github.com/user-attachments/assets/f9d4280a-5c7b-4f61-a938-f62cf2897442)


![image](https://github.com/user-attachments/assets/2ff21c39-567b-4597-9523-1bc1e7926950)

## Result:
The automation workflow successfully reads data from Sheet1 of the Excel file, processes it if required, and writes the data into Sheet2 (or another location) of the same Excel file. This confirms the ability of UiPath to interact with Excel for reading and writing operations.
