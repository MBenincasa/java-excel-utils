[![Maven Central](https://img.shields.io/maven-central/v/io.github.mbenincasa/java-excel-utils.svg?label=Maven%20Central)](https://search.maven.org/search?q=g:%22io.github.mbenincasa%22%20AND%20a:%22java-excel-utils%22)
[![GitHub release](https://img.shields.io/github/release/MBenincasa/java-excel-utils)](https://github.com/MBenincasa/java-excel-utils/releases/)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

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
    Workbook workbook = excelWorkbook.getWorkbook();
    Sheet sheet = excelSheet.getSheet();
    Row row = excelRow.getRow();
    Cell cell = excelCell.getCell();
}
```

One of the main features of the library is to be able to perform conversions. The **Converter** class has methods that convert **Excel <-> POJOs** and **Excel <-> CSV**.<br>
This is an example of Excel to POJOs:
```
public void ExcelToObjects() {
    File file = new File("./src/main/resources/car.xlsx");
    List<Car> cars = (List<Car>) Converter.excelToObjects(file, Car.class);
}
```

## Documentation
At the moment a real documentation is not yet available, but the javadocs and some examples are available in the `samples` package.<br>
Click [here](https://repo1.maven.org/maven2/io/github/mbenincasa/java-excel-utils/0.3.0/java-excel-utils-0.3.0-javadoc.jar) to download the javadocs.

## Minimum Requirements
Java 17 or above.

## Dependencies
- org.apache.poi:poi:jar:5.2.3
- org.apache.poi:poi-ooxml:jar:5.2.3
- org.projectlombok:lombok:jar:1.18.24
- commons-beanutils:commons-beanutils:jar:1.9.4
- com.opencsv:opencsv:jar:5.7.1
- org.junit.jupiter:junit-jupiter:jar:5.9.2
- org.junit.platform:junit-platform-suite-engine:jar:1.9.2

## Maven
```xml
<dependency>
  <groupId>io.github.mbenincasa</groupId>
  <artifactId>java-excel-utils</artifactId>
  <version>x.y.z</version>
</dependency>
```

## Roadmap
**Version 0.4.0** will bring new features to the Converter class. Current conversions will be reviewed. Conversions will be available that will also work with Stream I/O and byte[] in addition to the already present File.<p>
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
