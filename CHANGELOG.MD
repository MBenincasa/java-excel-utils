# v0.2.1
### Features
* New services for writing to an existing Excel file that take different input parameters
### Fixes
* Converting a value within a cell to an object returned null or failed. An error log is now written when an unsupported type is passed

# v0.2.0
### Deprecations
* All old tools are deprecated starting with this version. They have been replaced by new tools based on static methods.
### Features
* There are methods that allow you to convert a list of objects, or a CSV file, into an Excel file that already exists
* There are new utility methods such as returning the index of the last row or the last column of a row.
### Fixes
* The FileInputStream was not closed when the Workbook was opened from a File
### Enchantments
* Removed some checks and repetitive code

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
