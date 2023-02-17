package tools;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;
import org.apache.commons.io.FilenameUtils;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;

class ExcelUtilityTest {

    private final File excelFile = new File("./src/test/resources/employee.xlsx");

    @Test
    void countAllRowsOfAllSheets() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        List<Integer> results = ExcelUtility.countAllRowsOfAllSheets(excelFile);
        Assertions.assertEquals(3, results.get(0));
        Assertions.assertEquals(3, results.get(1));
    }

    @Test
    void testCountAllRowsOfAllSheets() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        List<Integer> results = ExcelUtility.countAllRowsOfAllSheets(excelFile, false);
        Assertions.assertEquals(3, results.get(0));
        Assertions.assertEquals(3, results.get(1));
    }

    @Test
    void countAllRows() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        Integer count = ExcelUtility.countAllRows(excelFile, "Office");
        Assertions.assertEquals(3, count);
    }

    @Test
    void testCountAllRows() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        Integer count = ExcelUtility.countAllRows(excelFile, "Office", false);
        Assertions.assertEquals(3, count);
    }

    @Test
    void checkExcelExtension() throws ExtensionNotValidException {
        String filename = excelFile.getName();
        String extension = ExcelUtility.checkExcelExtension(filename);
        Assertions.assertEquals("xlsx", extension);
    }

    @Test
    void isValidExcelExtension() {
        String filename = excelFile.getName();
        String extension = FilenameUtils.getExtension(filename);
        Assertions.assertEquals(true, ExcelUtility.isValidExcelExtension(extension));
    }

    @Test
    void getIndexLastRow() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        Integer index = ExcelUtility.getIndexLastRow(excelFile);
        Assertions.assertEquals(3, index);
    }

    @Test
    void testGetIndexLastRow() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        String sheetName = "Employee";
        Integer index = ExcelUtility.getIndexLastRow(excelFile, sheetName);
        Assertions.assertEquals(3, index);
    }

    @Test
    void getIndexLastColumn() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        Integer index = ExcelUtility.getIndexLastColumn(excelFile);
        Assertions.assertEquals(8, index);
    }

    @Test
    void testGetIndexLastColumn() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        String sheetName = "Employee";
        Integer index = ExcelUtility.getIndexLastColumn(excelFile, sheetName);
        Assertions.assertEquals(8, index);
    }

    @Test
    void testGetIndexLastColumn1() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        Integer index = ExcelUtility.getIndexLastColumn(excelFile, 1);
        Assertions.assertEquals(8, index);
    }

    @Test
    void testGetIndexLastColumn2() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        String sheetName = "Employee";
        Integer index = ExcelUtility.getIndexLastColumn(excelFile, sheetName, 1);
        Assertions.assertEquals(8, index);
    }
}