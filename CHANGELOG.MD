# v0.1.1
### Features
* A new method has been added which checks if an extension belongs to a Microsoft Excel file
### Fixes
* In the csvToExcel method, the CSVReader object was not closed at the end of the operation
* The objectsToExcel method did not check if the extension of the output file was correct

# v0.1.0
This is the first release.
### Features
* Converting a list of objects into an Excel file and vice versa
* Converting an Excel file to a CSV file and vice versa
* Utility methods which count all rows of one or more sheets and which check if an extension belongs to an Excel file
* Methods for opening, creating or closing a Workbook
* Methods to open, create or close a Sheet.
* Methods for retrieving a sheet or information relating to a sheet
