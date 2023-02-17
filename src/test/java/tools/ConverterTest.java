package tools;

import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvValidationException;
import enums.Extension;
import exceptions.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.*;
import tools.utils.Address;
import tools.utils.Person;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

class ConverterTest {

    private static final List<Person> persons = new ArrayList<>();
    private static final List<Address> addresses = new ArrayList<>();
    private static final File excelFile = new File("./src/test/resources/person.xlsx");
    private static final File csvFile = new File("./src/test/resources/person.csv");
    private static final File csvFile2 = new File("./src/test/resources/person_2.csv");

    @BeforeAll
    static void beforeAll() {
        persons.add(new Person("Rossi", "Mario", 20));
        addresses.add(new Address("Milano", "Corso Como, 4"));
    }

    @Test
    void objectsToExcel() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(persons, Person.class);
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void testObjectsToExcel() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(persons, Person.class, "person");
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void testObjectsToExcel1() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(persons, Person.class, "./src/", "person");
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void testObjectsToExcel2() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(persons, Person.class, "./src/", "person", false);
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void testObjectsToExcel3() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(persons, Person.class, false);
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void testObjectsToExcel4() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(persons, Person.class, "person", false);
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void testObjectsToExcel5() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(persons, Person.class, "./src/", "person", Extension.XLSX);
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void testObjectsToExcel6() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(persons, Person.class, Extension.XLSX);
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void testObjectsToExcel7() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(persons, Person.class, Extension.XLSX, false);
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void testObjectsToExcel8() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(persons, Person.class, "person", Extension.XLSX);
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void testObjectsToExcel9() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(persons, Person.class, "person", Extension.XLSX, false);
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void testObjectsToExcel10() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(persons, Person.class, "./src/", "person", Extension.XLSX, false);
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void objectsToExistingExcel() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(addresses, Address.class, false);
        Converter.objectsToExistingExcel(excelFile, persons, Person.class);
        Sheet sheet = SheetUtility.get(excelFile, "Person");
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void testObjectsToExistingExcel() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(addresses, Address.class, false);
        Converter.objectsToExistingExcel(excelFile, persons, Person.class, false);
        Sheet sheet = SheetUtility.get(excelFile, "Person");
        Row row = sheet.getRow(0);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        excelFile.delete();
    }

    @Test
    void testObjectsToExistingExcel1() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(addresses, Address.class, false);
        Workbook workbook = WorkbookUtility.open(excelFile);
        Converter.objectsToExistingExcel(workbook, persons, Person.class);
        FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
        workbook.write(fileOutputStream);
        Sheet sheet = SheetUtility.get(excelFile, "Person");
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        WorkbookUtility.close(workbook, fileOutputStream);
        excelFile.delete();
    }

    @Test
    void testObjectsToExistingExcel2() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, IllegalAccessException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.objectsToExcel(addresses, Address.class, false);
        Workbook workbook = WorkbookUtility.open(excelFile);
        Converter.objectsToExistingExcel(workbook, persons, Person.class, false);
        FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
        workbook.write(fileOutputStream);
        Sheet sheet = SheetUtility.get(excelFile, "Person");
        Row row = sheet.getRow(0);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals(20, row.getCell(2).getNumericCellValue());
        WorkbookUtility.close(workbook, fileOutputStream);
        excelFile.delete();
    }

    @Test
    void excelToObjects() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException, InvocationTargetException, HeaderNotPresentException, InstantiationException, IllegalAccessException, NoSuchMethodException {
        List<Person> persons = (List<Person>) Converter.excelToObjects(excelFile, Person.class);
        Assertions.assertEquals("Rossi", persons.get(0).getLastName());
        Assertions.assertEquals("Mario", persons.get(0).getName());
        Assertions.assertEquals(20, persons.get(0).getAge());
    }

    @Test
    void testExcelToObjects() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException, InvocationTargetException, HeaderNotPresentException, IllegalAccessException, NoSuchMethodException, InstantiationException {
        List<Person> persons = (List<Person>) Converter.excelToObjects(excelFile, Person.class, "Person");
        Assertions.assertEquals("Rossi", persons.get(0).getLastName());
        Assertions.assertEquals("Mario", persons.get(0).getName());
        Assertions.assertEquals(20, persons.get(0).getAge());
    }

    @Test
    void excelToCsv() throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException, CsvValidationException {
        File csvFile = Converter.excelToCsv(excelFile);
        FileReader fileReader = new FileReader(csvFile);
        CSVReader csvReader = new CSVReader(fileReader);
        String[] values = csvReader.readNext();
        Assertions.assertEquals("LAST NAME", values[0]);
        Assertions.assertEquals("NAME", values[1]);
        Assertions.assertEquals("AGE", values[2]);
        values = csvReader.readNext();
        Assertions.assertEquals("Rossi", values[0]);
        Assertions.assertEquals("Mario", values[1]);
        Assertions.assertEquals(20, Integer.parseInt(values[2]));
        csvReader.close();
        csvFile.delete();
    }

    @Test
    void testExcelToCsv() throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException, CsvValidationException {
        File csvFile = Converter.excelToCsv(excelFile, "Person");
        FileReader fileReader = new FileReader(csvFile);
        CSVReader csvReader = new CSVReader(fileReader);
        String[] values = csvReader.readNext();
        Assertions.assertEquals("LAST NAME", values[0]);
        Assertions.assertEquals("NAME", values[1]);
        Assertions.assertEquals("AGE", values[2]);
        values = csvReader.readNext();
        Assertions.assertEquals("Rossi", values[0]);
        Assertions.assertEquals("Mario", values[1]);
        Assertions.assertEquals(20, Integer.parseInt(values[2]));
        csvReader.close();
        csvFile.delete();
    }

    @Test
    void testExcelToCsv1() throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException, CsvValidationException {
        File csvFile = Converter.excelToCsv(excelFile, "./src/", "person");
        FileReader fileReader = new FileReader(csvFile);
        CSVReader csvReader = new CSVReader(fileReader);
        String[] values = csvReader.readNext();
        Assertions.assertEquals("LAST NAME", values[0]);
        Assertions.assertEquals("NAME", values[1]);
        Assertions.assertEquals("AGE", values[2]);
        values = csvReader.readNext();
        Assertions.assertEquals("Rossi", values[0]);
        Assertions.assertEquals("Mario", values[1]);
        Assertions.assertEquals(20, Integer.parseInt(values[2]));
        csvReader.close();
        csvFile.delete();
    }

    @Test
    void testExcelToCsv2() throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException, CsvValidationException {
        File csvFile = Converter.excelToCsv(excelFile, "./src/", "person", "Person");
        FileReader fileReader = new FileReader(csvFile);
        CSVReader csvReader = new CSVReader(fileReader);
        String[] values = csvReader.readNext();
        Assertions.assertEquals("LAST NAME", values[0]);
        Assertions.assertEquals("NAME", values[1]);
        Assertions.assertEquals("AGE", values[2]);
        values = csvReader.readNext();
        Assertions.assertEquals("Rossi", values[0]);
        Assertions.assertEquals("Mario", values[1]);
        Assertions.assertEquals(20, Integer.parseInt(values[2]));
        csvReader.close();
        csvFile.delete();
    }

    @Test
    void csvToExcel() throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.csvToExcel(csvFile);
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("20", row.getCell(2).getStringCellValue());
        excelFile.delete();
    }

    @Test
    void testCsvToExcel() throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.csvToExcel(csvFile, "person");
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("20", row.getCell(2).getStringCellValue());
        excelFile.delete();
    }

    @Test
    void testCsvToExcel1() throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.csvToExcel(csvFile, "./src/", "person");
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("20", row.getCell(2).getStringCellValue());
        excelFile.delete();
    }

    @Test
    void testCsvToExcel2() throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.csvToExcel(csvFile, "./src/", "person", Extension.XLSX);
        Sheet sheet = SheetUtility.get(excelFile);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("20", row.getCell(2).getStringCellValue());
        excelFile.delete();
    }

    @Test
    void csvToExistingExcel() throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.csvToExcel(csvFile);
        Converter.csvToExistingExcel(excelFile, csvFile2);
        Sheet sheet = SheetUtility.get(excelFile, 1);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("20", row.getCell(2).getStringCellValue());
        excelFile.delete();
    }

    @Test
    void testCsvToExistingExcel() throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.csvToExcel(csvFile);
        FileReader fileReader = new FileReader(csvFile2);
        CSVReader csvReader = new CSVReader(fileReader);
        Converter.csvToExistingExcel(excelFile, csvReader);
        Sheet sheet = SheetUtility.get(excelFile, 1);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("20", row.getCell(2).getStringCellValue());
        csvReader.close();
        excelFile.delete();
    }

    @Test
    void testCsvToExistingExcel1() throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.csvToExcel(csvFile);
        Workbook workbook = WorkbookUtility.open(excelFile);
        Converter.csvToExistingExcel(workbook, csvFile2);
        FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
        workbook.write(fileOutputStream);
        Sheet sheet = SheetUtility.get(excelFile, 1);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("20", row.getCell(2).getStringCellValue());
        WorkbookUtility.close(workbook, fileOutputStream);
        excelFile.delete();
    }

    @Test
    void testCsvToExistingExcel2() throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        File excelFile = Converter.csvToExcel(csvFile);
        Workbook workbook = WorkbookUtility.open(excelFile);
        FileReader fileReader = new FileReader(csvFile2);
        CSVReader csvReader = new CSVReader(fileReader);
        Converter.csvToExistingExcel(workbook, csvReader);
        FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
        workbook.write(fileOutputStream);
        Sheet sheet = SheetUtility.get(excelFile, 1);
        Row row = sheet.getRow(0);
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = sheet.getRow(1);
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("20", row.getCell(2).getStringCellValue());
        WorkbookUtility.close(workbook, fileOutputStream, csvReader);
        excelFile.delete();
    }
}