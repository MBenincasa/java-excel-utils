# v0.5.0
### Features
* Methods for writing a matrix of values starting from a source cell or in append and the method for deleting a matrix of cells has been added to Excel Sheet
* Wrapper classes can be initialized with static of() methods
* There are methods parseToObject() and parseToList() in ExcelSheet which allow, through defined rules, to assign the values of precise cells to the fields of an object.
### Fixes
* The readValue() method correctly returns the Date type
* The readValue(Class<?> type) method checks better if the return value is a date
### Removed
* Methods that have been deprecated since v0.4.0 have been removed
### Changes
* Constructor methods of wrapper classes are now private
* The writeValue() method uses the yyyy-MM-dd and yyyy-MM-dd HH:mm style to represent date and time

# v0.4.2
### Features
* Now there is a method to zip a list of files
* There is a method to read a cell name
* There is a method that returns the indexes of a cell starting from the cell name
### Fixes
* The readValue() method, in ExcelCell, did not always return the correct type

# v0.4.1
### Features
* There are getOrCreate methods for a Row or a Cell
* There are methods to remove a Sheet, a Row or a Cell
* There are methods to retrieve a Row or a Cell based on an index
* There are methods for writing and reading a list of values from a Row
### Fixes
* ExcelSheet's getIndex() method did not return the correct index after deleting a Sheet
### Changes
* The static method that opens an ExcelWorkbook via InputStream no longer takes the extension as input and is no longer checked if it is correct

# v0.4.0
### Deprecations
* All methods of the Converter class up to version 0.2.1 have been deprecated
### Features
* There are new methods that convert Excel <-> POJOs and Excel <-> CSV. Besides conversion between files, conversions between streams and byte arrays are now available. Data is no longer returned in a List but in a Stream
* Excel <-> Json conversion is now available
* The writeAndClose method is available in ExcelWorkbook
* ExcelCell provides new ways to read a cell.
### Fixes
* The error is handled better when trying to create a Sheet already present in the Workbook
* The readCell(Class<?> type) method did not always return the requested type
### Removed
* The WorkbookUtility and SheetUtility classes have been removed
* Several ExcelUtility methods have been removed

# v0.3.0
### Deprecations
* The WorkbookUtility and SheetUtility classes have been deprecated
* Several ExcelUtility methods have been deprecated
### Features
* The ExcelWorkbook, ExcelSheet, ExcelRow and ExceCell classes have been added. These classes wrap the Apache POI classes
### Fixes
* Several bugs have been fixed
### Removed
* Classes that have been deprecated since v0.2.0 have been removed

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
