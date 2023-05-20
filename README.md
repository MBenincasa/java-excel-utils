[![Maven Central](https://img.shields.io/maven-central/v/io.github.mbenincasa/java-excel-utils.svg?label=Maven%20Central)](https://search.maven.org/search?q=g:%22io.github.mbenincasa%22%20AND%20a:%22java-excel-utils%22)
[![GitHub release](https://img.shields.io/github/release/MBenincasa/java-excel-utils)](https://github.com/MBenincasa/java-excel-utils/releases/)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)<br>
[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.paypal.com/donate/?hosted_button_id=WXYAJVFZD82BJ)

# Java Excel Utils

## Introduction
Java Excel Utils is a Java library with the aim of making working with Microsoft Excel sheets easier. The library uses Apache POI components extending their potential.<br>

## Quickstart
There are wrapper classes for the main components of an Excel sheet: **ExcelWorkbook**, **ExcelSheet**, **ExcelRow** and **ExcelCell**.
```
public void modelTest() {
    ExcelWorkbook excelWorkbook = ExcelWorkbook.create(Extension.XLSX);
    ExcelSheet excelSheet = excelWorkbook.createSheet("TEST");
    ExcelRow excelRow = excelSheet.createRow(0);
    ExcelCell excelCell = excelRow.createCell(0);
    excelCell.writeValue("Rossi");
    String lastName = excelCell.readValue(String.class);
}
```

This code snippet shows how wrappat components can be instantiated. The last lines show how it is possible to write and read inside a cell.<br>
The classes have many other features, such as the ability to count the number of rows and columns present excluding, if you want, the empty ones.<p>
At any time you can retrieve the associated Apache POI components.

```
public void toPOI() {
    // Initialize the components
    Workbook workbook = excelWorkbook.getWorkbook();
    Sheet sheet = excelSheet.getSheet();
    Row row = excelRow.getRow();
    Cell cell = excelCell.getCell();
}

public void fromPOI() {
    // Initialize the components
    ExcelWorkbook excelWorkbook = ExcelWorkbook.of(workbook);
    ExcelSheet excelSheet = ExcelSheet.of(sheet);
    ExcelRow excelRow = ExcelRow.of(row);
    ExcelCell excelCell = ExcelCell.of(cell);
}
```

One of the main features of the library is to be able to perform conversions. The **Converter** class has methods that convert **EXCEL <-> POJOs**, **EXCEL <-> CSV** and **EXCEL <-> JSON**<br>
It is also possible to zip a list of files.<br>
This is an example of Excel to POJOs:
```
public void excelToObjects() {
    // Initialize List<ExcelToObject<?>> excelToObjects...
    File file = new File("./src/main/resources/car.xlsx");
    Map<String, Stream<?>> map = Converter.excelFileToObjects(file, excelToObjects);
    for (Map.Entry<String, Stream<?>> entry : map.entrySet()) {
        System.out.println("Sheet: " + entry.getKey());
        System.out.println("Data: " + entry.getValue().toList());
    }
}
```

This is an example of POJOs to Excel:
```
public void objectsToExcel() {
    // Initialize List<ObjectToExcel<?>> list...
    list.add(new ObjectToExcel<>("Employee", Employee.class, employeeStream));
    list.add(new ObjectToExcel<>("Office", Office.class, officeStream));
    File fileOutput = Converter.objectsToExcelFile(list, Extension.XLSX, "./src/main/resources/result", true);
}
```

ExcelSheet provides two methods for parsing the Sheet into an object or a list of objects.<br>
The advantage of these methods comes from the annotations and the mapping class that allow you to define the positions of the values of each field and the rules on how the various objects are positioned
```
public void parseSheet() {
   ExcelWorkbook excelWorkbook = ExcelWorkbook.open(file);
   ExcelSheet excelSheet = excelWorkbook.getSheet("DATA");
   Employee employee = excelSheet.parseToObject(Employee.class, "A1");
   ExcelListParserMapping mapping = new ExcelListParserMapping("A1", Direction.VERTICAL, 8);
   List<Employee> employees = excelSheet.parseToList(Employee.class, mapping);
}
```

ExcelCell provides generic methods for reading a cell.
```
public void readValue() {
    // Initialize ExcelCell excelCell...
    Integer intVal = excelCell.readValue(Integer.class);
    Double doubleVal = excelCell.readValue();
    String stringVal = excelCell.readValueAsString();
}
```

## Documentation
At the moment a real documentation is not yet available, but the javadocs and some examples are available in the `samples` package.<br>
Click [here](https://mbenincasa.github.io/java-excel-utils/) to view the java docs.

## Minimum Requirements
Java 17 or above.

## Dependencies
- org.apache.poi:poi:jar:5.2.3
- org.apache.poi:poi-ooxml:jar:5.2.3
- org.projectlombok:lombok:jar:1.18.26
- com.opencsv:opencsv:jar:5.7.1
- com.fasterxml.jackson.core:jackson-databind:jar:2.15.0
- org.apache.logging.log4j:log4j-core:jar:2.20.0
- org.junit.jupiter:junit-jupiter:jar:5.9.3
- org.junit.platform:junit-platform-suite-engine:jar:1.9.3

## Maven
```xml
<dependency>
  <groupId>io.github.mbenincasa</groupId>
  <artifactId>java-excel-utils</artifactId>
  <version>0.5.0</version>
</dependency>
```

## Roadmap
In the next updates there will be improvements on the wrapper classes that make up the Excel sheet.<p>
In the future, new features will come that have not yet been well defined.

## Contributing
Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks again!
1. Fork the Project
2. Create your Feature Branch
3. Commit your Changes
4. Push to the Branch
5. Open a Pull Request

## License
Distributed under the GNU General Public License v3.0. See `LICENSE.md` for more information.

## Contact
Mirko Benincasa - mirkobenincasa44@gmail.com

## Donations
Another way to support and contribute to the project is by sending a donation. The project will always be free and open source.<br>
Click [here](https://www.paypal.com/donate/?hosted_button_id=WXYAJVFZD82BJ) to make a donation on PayPal
