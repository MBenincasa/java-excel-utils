package model;

import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import enums.Extension;
import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import org.apache.commons.io.FilenameUtils;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.*;

public class ExcelWorkbookTest {

    private final File excelFile = new File("./src/test/resources/employee.xlsx");
    private final File csvFile = new File("./src/test/resources/employee.csv");

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

}