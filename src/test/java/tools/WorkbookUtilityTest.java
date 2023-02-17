package tools;

import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import enums.Extension;
import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.*;

class WorkbookUtilityTest {

    private final File excelFile = new File("./src/test/resources/employee.xlsx");
    private final File csvFile = new File("./src/test/resources/employee.csv");

    @Test
    void open() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        Workbook workbook = WorkbookUtility.open(excelFile);
        Assertions.assertNotNull(workbook);
    }

    @Test
    void testOpen() throws IOException, OpenWorkbookException, ExtensionNotValidException {
        String extension = FilenameUtils.getExtension(excelFile.getName());
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        Workbook workbook = WorkbookUtility.open(fileInputStream, extension);
        Assertions.assertNotNull(workbook);
        fileInputStream.close();
    }

    @Test
    void create() {
        Workbook workbook = WorkbookUtility.create();
        Assertions.assertNotNull(workbook);
    }

    @Test
    void testCreate() throws ExtensionNotValidException {
        String extension = "xlsx";
        Workbook workbook = WorkbookUtility.create(extension);
        Assertions.assertNotNull(workbook);
    }

    @Test
    void testCreate1() {
        Workbook workbook = WorkbookUtility.create(Extension.XLSX);
        Assertions.assertNotNull(workbook);
    }

    @Test
    void close() {
        Workbook workbook = WorkbookUtility.create(Extension.XLSX);
        Assertions.assertDoesNotThrow(() -> WorkbookUtility.close(workbook));
    }

    @Test
    void testClose() throws FileNotFoundException {
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        Workbook workbook = WorkbookUtility.create();
        Assertions.assertDoesNotThrow(() -> WorkbookUtility.close(workbook, fileInputStream));
    }

    @Test
    void testClose1() throws FileNotFoundException {
        FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
        Workbook workbook = WorkbookUtility.create();
        Assertions.assertDoesNotThrow(() -> WorkbookUtility.close(workbook, fileOutputStream));
    }

    @Test
    void testClose2() throws FileNotFoundException {
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
        Workbook workbook = WorkbookUtility.create();
        Assertions.assertDoesNotThrow(() -> WorkbookUtility.close(workbook, fileOutputStream, fileInputStream));
    }

    @Test
    void testClose3() throws IOException {
        Workbook workbook = WorkbookUtility.create();
        FileWriter fileWriter = new FileWriter(csvFile);
        CSVWriter csvWriter = new CSVWriter(fileWriter);
        Assertions.assertDoesNotThrow(() -> WorkbookUtility.close(workbook, csvWriter));
    }

    @Test
    void testClose4() throws FileNotFoundException {
        Workbook workbook = WorkbookUtility.create();
        FileReader fileReader = new FileReader(csvFile);
        CSVReader csvReader = new CSVReader(fileReader);
        FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
        Assertions.assertDoesNotThrow(() -> WorkbookUtility.close(workbook, fileOutputStream, csvReader));

    }
}