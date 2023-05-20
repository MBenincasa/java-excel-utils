package io.github.mbenincasa.javaexcelutils.tools;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvValidationException;
import io.github.mbenincasa.javaexcelutils.enums.Extension;
import io.github.mbenincasa.javaexcelutils.exceptions.*;
import io.github.mbenincasa.javaexcelutils.model.converter.ExcelToObject;
import io.github.mbenincasa.javaexcelutils.model.converter.JsonToExcel;
import io.github.mbenincasa.javaexcelutils.model.converter.ObjectToExcel;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelCell;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelRow;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelSheet;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelWorkbook;
import io.github.mbenincasa.javaexcelutils.tools.utils.Office;
import org.apache.poi.ss.usermodel.Row;
import org.junit.jupiter.api.*;
import io.github.mbenincasa.javaexcelutils.tools.utils.Address;
import io.github.mbenincasa.javaexcelutils.tools.utils.Person;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

public class ConverterTest {

    private static final List<Person> persons = new ArrayList<>();
    private static final List<Address> addresses = new ArrayList<>();
    private static final File jsonFile = new File("./src/test/resources/office.json");
    private static final File excelFile = new File("./src/test/resources/person.xlsx");
    private static final File csvFile = new File("./src/test/resources/person.csv");
    private static final File csvFile2 = new File("./src/test/resources/person_2.csv");

    @BeforeAll
    static void beforeAll() {
        persons.add(new Person("Rossi", "Mario", 20));
        addresses.add(new Address("Milano", "Corso Como, 4"));
    }

    @Test
    void objectsToExcelFile() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException, ReadValueException, SheetAlreadyExistsException {
        Stream<Person> personStream = persons.stream();
        Stream<Address> addressStream = addresses.stream();
        List<ObjectToExcel<?>> list = new ArrayList<>();
        list.add(new ObjectToExcel<>("Person", Person.class, personStream));
        list.add(new ObjectToExcel<>("Address", Address.class, addressStream));
        File fileOutput = Converter.objectsToExcelFile(list, Extension.XLSX, "./src/test/resources/result", true);

        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(fileOutput);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        Assertions.assertEquals("Person", excelSheet.getName());
        ExcelRow excelRow = excelSheet.getRows().get(1);
        List<ExcelCell> excelCells = excelRow.getCells();
        Assertions.assertEquals("Rossi", excelCells.get(0).readValue(String.class));
        Assertions.assertEquals("Mario", excelCells.get(1).readValue(String.class));
        Assertions.assertEquals(20, excelCells.get(2).readValue(Integer.class));
        ExcelSheet excelSheet1 = excelWorkbook.getSheet(1);
        Assertions.assertEquals("Address", excelSheet1.getName());
        ExcelRow excelRow1 = excelSheet1.getRows().get(1);
        List<ExcelCell> excelCells1 = excelRow1.getCells();
        Assertions.assertEquals("Milano", excelCells1.get(0).readValue(String.class));
        Assertions.assertEquals("Corso Como, 4", excelCells1.get(1).readValue(String.class));
        fileOutput.delete();
    }

    @Test
    void objectsToExcelByte() throws ExtensionNotValidException, IOException, SheetNotFoundException, ReadValueException, OpenWorkbookException, SheetAlreadyExistsException {
        Stream<Person> personStream = persons.stream();
        Stream<Address> addressStream = addresses.stream();
        List<ObjectToExcel<?>> list = new ArrayList<>();
        list.add(new ObjectToExcel<>("Person", Person.class, personStream));
        list.add(new ObjectToExcel<>("Address", Address.class, addressStream));
        byte[] bytes = Converter.objectsToExcelByte(list, Extension.XLSX, true);
        File fileOutput = new File("./src/test/resources/result.xlsx");
        FileOutputStream fileOutputStream = new FileOutputStream(fileOutput);
        fileOutputStream.write(bytes);
        fileOutputStream.close();

        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(fileOutput);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        Assertions.assertEquals("Person", excelSheet.getName());
        ExcelRow excelRow = excelSheet.getRows().get(1);
        List<ExcelCell> excelCells = excelRow.getCells();
        Assertions.assertEquals("Rossi", excelCells.get(0).readValue(String.class));
        Assertions.assertEquals("Mario", excelCells.get(1).readValue(String.class));
        Assertions.assertEquals(20, excelCells.get(2).readValue(Integer.class));
        ExcelSheet excelSheet1 = excelWorkbook.getSheet(1);
        Assertions.assertEquals("Address", excelSheet1.getName());
        ExcelRow excelRow1 = excelSheet1.getRows().get(1);
        List<ExcelCell> excelCells1 = excelRow1.getCells();
        Assertions.assertEquals("Milano", excelCells1.get(0).readValue(String.class));
        Assertions.assertEquals("Corso Como, 4", excelCells1.get(1).readValue(String.class));
        fileOutput.delete();
    }

    @Test
    void objectsToExcelStream() throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException, ReadValueException, SheetAlreadyExistsException {
        Stream<Person> personStream = persons.stream();
        Stream<Address> addressStream = addresses.stream();
        List<ObjectToExcel<?>> list = new ArrayList<>();
        list.add(new ObjectToExcel<>("Person", Person.class, personStream));
        list.add(new ObjectToExcel<>("Address", Address.class, addressStream));
        ByteArrayOutputStream outputStream = Converter.objectsToExcelStream(list, Extension.XLSX, true);
        byte[] bytes = outputStream.toByteArray();
        File fileOutput = new File("./src/test/resources/result.xlsx");
        FileOutputStream fileOutputStream = new FileOutputStream(fileOutput);
        fileOutputStream.write(bytes);
        fileOutputStream.close();

        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(fileOutput);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        Assertions.assertEquals("Person", excelSheet.getName());
        ExcelRow excelRow = excelSheet.getRows().get(1);
        List<ExcelCell> excelCells = excelRow.getCells();
        Assertions.assertEquals("Rossi", excelCells.get(0).readValue(String.class));
        Assertions.assertEquals("Mario", excelCells.get(1).readValue(String.class));
        Assertions.assertEquals(20, excelCells.get(2).readValue(Integer.class));
        ExcelSheet excelSheet1 = excelWorkbook.getSheet(1);
        Assertions.assertEquals("Address", excelSheet1.getName());
        ExcelRow excelRow1 = excelSheet1.getRows().get(1);
        List<ExcelCell> excelCells1 = excelRow1.getCells();
        Assertions.assertEquals("Milano", excelCells1.get(0).readValue(String.class));
        Assertions.assertEquals("Corso Como, 4", excelCells1.get(1).readValue(String.class));
        fileOutput.delete();
    }

    @Test
    void excelByteToObjects() throws IOException, OpenWorkbookException, SheetNotFoundException, ReadValueException, HeaderNotPresentException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        ExcelToObject<Person> personExcelToObject = new ExcelToObject<>("Person", Person.class);
        List<ExcelToObject<?>> excelToObjects = new ArrayList<>();
        excelToObjects.add(personExcelToObject);
        byte[] bytes = Files.readAllBytes(excelFile.toPath());
        Map<String, Stream<?>> map = Converter.excelByteToObjects(bytes, excelToObjects);
        List<Person> people = (List<Person>) map.get("Person").toList();
        Assertions.assertEquals("Rossi", people.get(0).getLastName());
        Assertions.assertEquals("Mario", people.get(0).getName());
        Assertions.assertEquals(20, people.get(0).getAge());
    }

    @Test
    void excelFileToObjects() throws OpenWorkbookException, SheetNotFoundException, ReadValueException, IOException, HeaderNotPresentException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException, ExtensionNotValidException {
        ExcelToObject<Person> personExcelToObject = new ExcelToObject<>("Person", Person.class);
        List<ExcelToObject<?>> excelToObjects = new ArrayList<>();
        excelToObjects.add(personExcelToObject);
        Map<String, Stream<?>> map = Converter.excelFileToObjects(excelFile, excelToObjects);
        List<Person> people = (List<Person>) map.get("Person").toList();
        Assertions.assertEquals("Rossi", people.get(0).getLastName());
        Assertions.assertEquals("Mario", people.get(0).getName());
        Assertions.assertEquals(20, people.get(0).getAge());
    }

    @Test
    void excelStreamToObjects() throws IOException, OpenWorkbookException, SheetNotFoundException, ReadValueException, HeaderNotPresentException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        ExcelToObject<Person> personExcelToObject = new ExcelToObject<>("Person", Person.class);
        List<ExcelToObject<?>> excelToObjects = new ArrayList<>();
        excelToObjects.add(personExcelToObject);
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        Map<String, Stream<?>> map = Converter.excelStreamToObjects(fileInputStream, excelToObjects);
        List<Person> people = (List<Person>) map.get("Person").toList();
        Assertions.assertEquals("Rossi", people.get(0).getLastName());
        Assertions.assertEquals("Mario", people.get(0).getName());
        Assertions.assertEquals(20, people.get(0).getAge());
        fileInputStream.close();
    }

    @Test
    void excelToCsvByte() throws IOException, OpenWorkbookException, CsvValidationException {
        byte[] bytes = Files.readAllBytes(excelFile.toPath());
        Map<String, byte[]> byteMap = Converter.excelToCsvByte(bytes);
        File csvFile = new File("person.csv");
        FileOutputStream fileOutputStream = new FileOutputStream(csvFile);
        fileOutputStream.write(byteMap.get("Person"));
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
        fileOutputStream.close();
        csvReader.close();
        csvFile.delete();
    }

    @Test
    void excelToCsvFile() throws OpenWorkbookException, IOException, CsvValidationException, ExtensionNotValidException {
        Map<String, File> fileMap = Converter.excelToCsvFile(excelFile, "./src/");
        FileReader fileReader = new FileReader(fileMap.get("Person"));
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
        fileMap.get("Person").delete();
    }

    @Test
    void excelToCsvStream() throws IOException, OpenWorkbookException, CsvValidationException {
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        Map<String, ByteArrayOutputStream> outputStreamMap = Converter.excelToCsvStream(fileInputStream);
        FileOutputStream fileOutputStream = new FileOutputStream("person.csv");
        ByteArrayOutputStream baos = outputStreamMap.get("Person");
        fileOutputStream.write(baos.toByteArray());
        File csvFile = new File("person.csv");
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
        fileOutputStream.close();
        csvReader.close();
        csvFile.delete();
        fileInputStream.close();
    }

    @Test
    void csvToExcelByte() throws IOException, CsvValidationException, ExtensionNotValidException, OpenWorkbookException, SheetNotFoundException, SheetAlreadyExistsException {
        byte[] bytes = Files.readAllBytes(csvFile.toPath());
        byte[] bytesResult = Converter.csvToExcelByte(bytes, "Test", Extension.XLSX);
        File excelFile = new File("./test.xlsx");
        FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
        fileOutputStream.write(bytesResult);
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet("Test");
        ExcelRow excelRow = excelSheet.getRows().get(0);
        Row row = excelRow.getRow();
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = excelSheet.getRows().get(1).getRow();
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("20", row.getCell(2).getStringCellValue());
        fileOutputStream.close();
        excelFile.delete();
    }

    @Test
    void csvToExcelFile() throws CsvValidationException, ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException, SheetAlreadyExistsException {
        File excelFile = Converter.csvToExcelFile(csvFile, "Test", "./test", Extension.XLSX);
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet("Test");
        ExcelRow excelRow = excelSheet.getRows().get(0);
        Row row = excelRow.getRow();
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = excelSheet.getRows().get(1).getRow();
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("20", row.getCell(2).getStringCellValue());
        excelFile.delete();
    }

    @Test
    void csvToExcelStream() throws IOException, CsvValidationException, ExtensionNotValidException, OpenWorkbookException, SheetNotFoundException, SheetAlreadyExistsException {
        FileInputStream fileInputStream = new FileInputStream(csvFile);
        ByteArrayOutputStream baos = Converter.csvToExcelStream(fileInputStream, "Test", Extension.XLSX);
        FileOutputStream fileOutputStream = new FileOutputStream("./test.xlsx");
        fileOutputStream.write(baos.toByteArray());
        File excelFile = new File("./test.xlsx");
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet("Test");
        ExcelRow excelRow = excelSheet.getRows().get(0);
        Row row = excelRow.getRow();
        Assertions.assertEquals("LAST NAME", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("NAME", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("AGE", row.getCell(2).getStringCellValue());
        row = excelSheet.getRows().get(1).getRow();
        Assertions.assertEquals("Rossi", row.getCell(0).getStringCellValue());
        Assertions.assertEquals("Mario", row.getCell(1).getStringCellValue());
        Assertions.assertEquals("20", row.getCell(2).getStringCellValue());
        fileOutputStream.close();
        excelFile.delete();
    }

    @Test
    void excelToJsonByte() throws IOException, OpenWorkbookException {
        byte[] bytes = Files.readAllBytes(excelFile.toPath());
        byte[] byteJson = Converter.excelToJsonByte(bytes);
        File jsonFile = new File("./src/test/resources/test.json");
        FileOutputStream fileOutputStream = new FileOutputStream(jsonFile);
        fileOutputStream.write(byteJson);
        FileInputStream fileInputStream = new FileInputStream(jsonFile);
        ObjectMapper objectMapper = new ObjectMapper();
        JsonNode jsonNode = objectMapper.readTree(fileInputStream);
        Assertions.assertEquals("LAST NAME", jsonNode.get("Person").get("row_1").get("col_1").asText());
        Assertions.assertEquals("NAME", jsonNode.get("Person").get("row_1").get("col_2").asText());
        Assertions.assertEquals("AGE", jsonNode.get("Person").get("row_1").get("col_3").asText());
        Assertions.assertEquals("Rossi", jsonNode.get("Person").get("row_2").get("col_1").asText());
        Assertions.assertEquals("Mario", jsonNode.get("Person").get("row_2").get("col_2").asText());
        Assertions.assertEquals("20", jsonNode.get("Person").get("row_2").get("col_3").asText());
        fileInputStream.close();
        fileOutputStream.close();
        jsonFile.delete();
    }

    @Test
    void excelToJsonFile() throws OpenWorkbookException, IOException, ExtensionNotValidException {
        File jsonFile = Converter.excelToJsonFile(excelFile, "./result");
        FileInputStream fileInputStream = new FileInputStream(jsonFile);
        ObjectMapper objectMapper = new ObjectMapper();
        JsonNode jsonNode = objectMapper.readTree(fileInputStream);
        Assertions.assertEquals("LAST NAME", jsonNode.get("Person").get("row_1").get("col_1").asText());
        Assertions.assertEquals("NAME", jsonNode.get("Person").get("row_1").get("col_2").asText());
        Assertions.assertEquals("AGE", jsonNode.get("Person").get("row_1").get("col_3").asText());
        Assertions.assertEquals("Rossi", jsonNode.get("Person").get("row_2").get("col_1").asText());
        Assertions.assertEquals("Mario", jsonNode.get("Person").get("row_2").get("col_2").asText());
        Assertions.assertEquals("20", jsonNode.get("Person").get("row_2").get("col_3").asText());
        fileInputStream.close();
        jsonFile.delete();
    }

    @Test
    void excelToJsonStream() throws IOException, OpenWorkbookException {
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        ByteArrayOutputStream baos = Converter.excelToJsonStream(fileInputStream);
        FileOutputStream fileOutputStream = new FileOutputStream("./src/test/resources/test.json");
        fileOutputStream.write(baos.toByteArray());
        File jsonFile = new File("./src/test/resources/test.json");
        ObjectMapper objectMapper = new ObjectMapper();
        JsonNode jsonNode = objectMapper.readTree(jsonFile);
        Assertions.assertEquals("LAST NAME", jsonNode.get("Person").get("row_1").get("col_1").asText());
        Assertions.assertEquals("NAME", jsonNode.get("Person").get("row_1").get("col_2").asText());
        Assertions.assertEquals("AGE", jsonNode.get("Person").get("row_1").get("col_3").asText());
        Assertions.assertEquals("Rossi", jsonNode.get("Person").get("row_2").get("col_1").asText());
        Assertions.assertEquals("Mario", jsonNode.get("Person").get("row_2").get("col_2").asText());
        Assertions.assertEquals("20", jsonNode.get("Person").get("row_2").get("col_3").asText());
        fileInputStream.close();
        fileOutputStream.close();
        jsonFile.delete();
    }

    @Test
    void jsonToExcelByte() throws IOException, ExtensionNotValidException, SheetAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ReadValueException {
        JsonToExcel<Office> officeJsonToExcel = new JsonToExcel<>("office", Office.class);
        byte[] bytes = Files.readAllBytes(jsonFile.toPath());
        byte[] byteResult = Converter.jsonToExcelByte(bytes, officeJsonToExcel, Extension.XLSX, true);
        File excelFile = new File("./src/test/resourcestest.xlsx");
        FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
        fileOutputStream.write(byteResult);
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelRow excelRow = excelSheet.getRows().get(0);
        Assertions.assertEquals("CITY", excelRow.getCells().get(0).readValue(String.class));
        Assertions.assertEquals("PROVINCE", excelRow.getCells().get(1).readValue(String.class));
        Assertions.assertEquals("NUMBER OF STATIONS", excelRow.getCells().get(2).readValue(String.class));
        ExcelRow excelRow1 = excelSheet.getRows().get(1);
        Assertions.assertEquals("Nocera Inferiore", excelRow1.getCells().get(0).readValue(String.class));
        Assertions.assertEquals("Salerno", excelRow1.getCells().get(1).readValue(String.class));
        Assertions.assertEquals(21, excelRow1.getCells().get(2).readValue(Integer.class));
        fileOutputStream.close();
        excelFile.delete();

    }

    @Test
    void jsonToExcelFile() throws FileAlreadyExistsException, ExtensionNotValidException, IOException, SheetAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ReadValueException {
        JsonToExcel<Office> officeJsonToExcel = new JsonToExcel<>("office", Office.class);
        File excelFile = Converter.jsonToExcelFile(jsonFile, officeJsonToExcel, Extension.XLSX, "./excel", true);
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelRow excelRow = excelSheet.getRows().get(0);
        Assertions.assertEquals("CITY", excelRow.getCells().get(0).readValue(String.class));
        Assertions.assertEquals("PROVINCE", excelRow.getCells().get(1).readValue(String.class));
        Assertions.assertEquals("NUMBER OF STATIONS", excelRow.getCells().get(2).readValue(String.class));
        ExcelRow excelRow1 = excelSheet.getRows().get(1);
        Assertions.assertEquals("Nocera Inferiore", excelRow1.getCells().get(0).readValue(String.class));
        Assertions.assertEquals("Salerno", excelRow1.getCells().get(1).readValue(String.class));
        Assertions.assertEquals(21, excelRow1.getCells().get(2).readValue(Integer.class));
        excelFile.delete();
    }

    @Test
    void jsonToExcelStream() throws IOException, ExtensionNotValidException, SheetAlreadyExistsException, OpenWorkbookException, ReadValueException, SheetNotFoundException {
        FileInputStream fileInputStream = new FileInputStream(jsonFile);
        JsonToExcel<Office> officeJsonToExcel = new JsonToExcel<>("office", Office.class);
        ByteArrayOutputStream baos = Converter.jsonToExcelStream(fileInputStream, officeJsonToExcel, Extension.XLSX, true);
        FileOutputStream fileOutputStream = new FileOutputStream("./src/test/resources/test.xlsx");
        fileOutputStream.write(baos.toByteArray());
        File excelFile = new File("./src/test/resources/test.xlsx");
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelRow excelRow = excelSheet.getRows().get(0);
        Assertions.assertEquals("CITY", excelRow.getCells().get(0).readValue(String.class));
        Assertions.assertEquals("PROVINCE", excelRow.getCells().get(1).readValue(String.class));
        Assertions.assertEquals("NUMBER OF STATIONS", excelRow.getCells().get(2).readValue(String.class));
        ExcelRow excelRow1 = excelSheet.getRows().get(1);
        Assertions.assertEquals("Nocera Inferiore", excelRow1.getCells().get(0).readValue(String.class));
        Assertions.assertEquals("Salerno", excelRow1.getCells().get(1).readValue(String.class));
        Assertions.assertEquals(21, excelRow1.getCells().get(2).readValue(Integer.class));
        fileInputStream.close();
        fileOutputStream.close();
        excelFile.delete();
    }
}