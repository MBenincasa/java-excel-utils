package io.github.mbenincasa.javaexcelutils.model.excel;

import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import io.github.mbenincasa.javaexcelutils.enums.Extension;
import io.github.mbenincasa.javaexcelutils.exceptions.ExtensionNotValidException;
import io.github.mbenincasa.javaexcelutils.exceptions.OpenWorkbookException;
import io.github.mbenincasa.javaexcelutils.exceptions.SheetAlreadyExistsException;
import io.github.mbenincasa.javaexcelutils.exceptions.SheetNotFoundException;
import io.github.mbenincasa.javaexcelutils.model.converter.ObjectToExcel;
import io.github.mbenincasa.javaexcelutils.tools.Converter;
import io.github.mbenincasa.javaexcelutils.tools.utils.Address;
import io.github.mbenincasa.javaexcelutils.tools.utils.Person;
import org.apache.commons.io.FilenameUtils;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

public class ExcelWorkbookTest {

    private final File excelFile = new File("./src/test/resources/employee.xlsx");
    private final File csvFile = new File("./src/test/resources/employee.csv");

    private static final List<Person> persons = new ArrayList<>();
    private static final List<Address> addresses = new ArrayList<>();

    @BeforeAll
    static void beforeAll() {
        persons.add(new Person("Rossi", "Mario", 20));
        addresses.add(new Address("Milano", "Corso Como, 4"));
    }

    @Test
    void open() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        Assertions.assertNotNull(excelWorkbook);
        Assertions.assertNotNull(excelWorkbook.getWorkbook());
    }

    @Test
    void testOpen() throws IOException, OpenWorkbookException, ExtensionNotValidException {
        String extension = FilenameUtils.getExtension(excelFile.getName());
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(fileInputStream, extension);
        Assertions.assertNotNull(excelWorkbook);
        Assertions.assertNotNull(excelWorkbook.getWorkbook());
    }

    @Test
    void create() {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.create();
        Assertions.assertNotNull(excelWorkbook);
        Assertions.assertNotNull(excelWorkbook.getWorkbook());
    }

    @Test
    void testCreate() throws ExtensionNotValidException {
        String extension = "xlsx";
        ExcelWorkbook excelWorkbook = ExcelWorkbook.create(extension);
        Assertions.assertNotNull(excelWorkbook);
        Assertions.assertNotNull(excelWorkbook.getWorkbook());
    }

    @Test
    void testCreate1() {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.create(Extension.XLSX);
        Assertions.assertNotNull(excelWorkbook);
        Assertions.assertNotNull(excelWorkbook.getWorkbook());
    }

    @Test
    void close() {
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(Extension.XLSX);
        Assertions.assertDoesNotThrow(() -> excelWorkbook.close());
    }

    @Test
    void testClose() throws FileNotFoundException {
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(Extension.XLSX);
        Assertions.assertDoesNotThrow(() -> excelWorkbook.close(fileInputStream));
    }

    @Test
    void testClose1() throws FileNotFoundException {
        FileOutputStream fileOutputStream = new FileOutputStream(excelFile, true);
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(Extension.XLSX);
        Assertions.assertDoesNotThrow(() -> excelWorkbook.close(fileOutputStream));
    }

    @Test
    void testClose2() throws FileNotFoundException {
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        FileOutputStream fileOutputStream = new FileOutputStream(excelFile, true);
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(Extension.XLSX);
        Assertions.assertDoesNotThrow(() -> excelWorkbook.close(fileOutputStream, fileInputStream));
    }

    @Test
    void testClose3() throws IOException {
        FileWriter fileWriter = new FileWriter(csvFile, true);
        CSVWriter csvWriter = new CSVWriter(fileWriter);
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(Extension.XLSX);
        Assertions.assertDoesNotThrow(() -> excelWorkbook.close(csvWriter));
    }

    @Test
    void testClose4() throws FileNotFoundException {
        FileReader fileReader = new FileReader(csvFile);
        CSVReader csvReader = new CSVReader(fileReader);
        FileOutputStream fileOutputStream = new FileOutputStream(excelFile, true);
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(Extension.XLSX);
        Assertions.assertDoesNotThrow(() -> excelWorkbook.close(fileOutputStream, csvReader));
    }

    @Test
    void countSheets() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        Assertions.assertEquals(2, excelWorkbook.countSheets());
    }

    @Test
    void createSheet() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetAlreadyExistsException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.createSheet("Test");
        Assertions.assertNotNull(excelSheet.getSheet());
        Assertions.assertEquals("Test", excelSheet.getName());
    }

    @Test
    void testCreateSheet() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.createSheet();
        Assertions.assertNotNull(excelSheet.getSheet());
    }

    @Test
    void getSheets() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        List<ExcelSheet> excelSheets = excelWorkbook.getSheets();
        Assertions.assertEquals("Employee", excelSheets.get(0).getName());
        Assertions.assertEquals("Office", excelSheets.get(1).getName());
    }

    @Test
    void getSheet() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        Assertions.assertEquals("Employee", excelWorkbook.getSheet(0).getName());
    }

    @Test
    void testGetSheet() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        Assertions.assertEquals("Employee", excelWorkbook.getSheet("Employee").getName());
    }

    @Test
    void testGetSheet1() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        Assertions.assertEquals("Employee", excelWorkbook.getSheet().getName());
    }

    @Test
    void getSheetOrCreate() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        Assertions.assertEquals("Employee_2", excelWorkbook.getSheetOrCreate("Employee_2").getName());
    }

    @Test
    void isSheetPresent() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        Assertions.assertEquals(true, excelWorkbook.isSheetPresent("Employee"));
    }

    @Test
    void testIsSheetPresent() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        Assertions.assertEquals(false, excelWorkbook.isSheetPresent(3));
    }

    @Test
    void isSheetNull() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        Assertions.assertEquals(false, excelWorkbook.isSheetNull("Office"));
    }

    @Test
    void testIsSheetNull() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        Assertions.assertEquals(true, excelWorkbook.isSheetNull(3));
    }

    @Test
    void getFormulaEvaluator() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        Assertions.assertNotNull(excelWorkbook.getFormulaEvaluator());
    }

    @Test
    void writeAndClose() throws ExtensionNotValidException, IOException, SheetAlreadyExistsException {
        Stream<Person> personStream = persons.stream();
        Stream<Address> addressStream = addresses.stream();
        List<ObjectToExcel<?>> list = new ArrayList<>();
        list.add(new ObjectToExcel<>("Person", Person.class, personStream));
        list.add(new ObjectToExcel<>("Address", Address.class, addressStream));
        ByteArrayOutputStream outputStream = (ByteArrayOutputStream) Converter.objectsToExcelStream(list, Extension.XLSX, true);
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(Extension.XLSX);
        Assertions.assertDoesNotThrow(() -> excelWorkbook.writeAndClose(outputStream));
    }

    @Test
    void removeSheet() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelSheet excelSheet1 = excelWorkbook.getSheet(1);
        Assertions.assertDoesNotThrow(() -> excelWorkbook.removeSheet(excelSheet.getIndex()));
        Assertions.assertEquals(0, excelSheet1.getIndex());
    }
}