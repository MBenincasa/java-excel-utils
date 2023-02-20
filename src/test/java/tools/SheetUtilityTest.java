package tools;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class SheetUtilityTest {

    private final File excelFile = new File("./src/test/resources/employee.xlsx");

    @Test
    void length() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        Integer length = SheetUtility.length(excelFile);
        Assertions.assertEquals(length, 2);
    }

    @Test
    void getNames() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        List<String> names = SheetUtility.getNames(excelFile);
        Assertions.assertEquals(names.get(0), "Employee");
        Assertions.assertEquals(names.get(1), "Office");

    }

    @Test
    void getIndex() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        Integer index = SheetUtility.getIndex(excelFile, "Employee");
        Assertions.assertEquals(index, 0);
    }

    @Test
    void getName() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        String name = SheetUtility.getName(excelFile, 0);
        Assertions.assertEquals(name, "Employee");
    }

    @Test
    void create() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        Sheet sheet = SheetUtility.create(excelFile);
        Assertions.assertNotNull(sheet);
    }

    @Test
    void testCreate() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        String sheetName = "Admin";
        Sheet sheet = SheetUtility.create(excelFile, sheetName);
        Assertions.assertNotNull(sheet);
        Assertions.assertEquals(sheet.getSheetName(), sheetName);
    }

    @Test
    void testCreate1() throws ExtensionNotValidException {
        String extension = "xlsx";
        Workbook workbook = WorkbookUtility.create(extension);
        Sheet sheet = SheetUtility.create(workbook);
        Assertions.assertNotNull(sheet);
    }

    @Test
    void testCreate2() throws ExtensionNotValidException {
        String extension = "xlsx";
        String sheetName = "Admin";
        Workbook workbook = WorkbookUtility.create(extension);
        Sheet sheet = SheetUtility.create(workbook, sheetName);
        Assertions.assertNotNull(sheet);
        Assertions.assertEquals(sheet.getSheetName(), sheetName);
    }

    @Test
    void get() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        Sheet sheet = SheetUtility.get(excelFile);
        Assertions.assertNotNull(sheet);
    }

    @Test
    void testGet() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        String sheetName = "Employee";
        Sheet sheet = SheetUtility.get(excelFile, sheetName);
        Assertions.assertNotNull(sheet);
        Assertions.assertEquals(sheet.getSheetName(), sheetName);
    }

    @Test
    void testGet1() throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        String sheetName = "Employee";
        Sheet sheet = SheetUtility.get(excelFile, 0);
        Assertions.assertNotNull(sheet);
        Assertions.assertEquals(sheet.getSheetName(), sheetName);
    }

    @Test
    void testGet2() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        Workbook workbook = WorkbookUtility.open(excelFile);
        Sheet sheet = SheetUtility.get(workbook);
        Assertions.assertNotNull(sheet);
    }

    @Test
    void testGet3() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        String sheetName = "Employee";
        Workbook workbook = WorkbookUtility.open(excelFile);
        Sheet sheet = SheetUtility.get(workbook, sheetName);
        Assertions.assertNotNull(sheet);
        Assertions.assertEquals(sheet.getSheetName(), sheetName);
    }

    @Test
    void testGet4() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        String sheetName = "Employee";
        Workbook workbook = WorkbookUtility.open(excelFile);
        Sheet sheet = SheetUtility.get(workbook, 0);
        Assertions.assertNotNull(sheet);
        Assertions.assertEquals(sheet.getSheetName(), sheetName);
    }

    @Test
    void getOrCreate() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        String sheetName = "Admin";
        Workbook workbook = WorkbookUtility.open(excelFile);
        Sheet sheet = SheetUtility.getOrCreate(workbook, sheetName);
        Assertions.assertNotNull(sheet);
        Assertions.assertEquals(sheet.getSheetName(), sheetName);
    }

    @Test
    void isPresent() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        String sheetName = "Employee";
        Workbook workbook = WorkbookUtility.open(excelFile);
        Assertions.assertEquals(true, SheetUtility.isPresent(workbook, sheetName));
    }

    @Test
    void testIsPresent() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        Workbook workbook = WorkbookUtility.open(excelFile);
        Assertions.assertEquals(true, SheetUtility.isPresent(workbook, 1));
    }

    @Test
    void isNull() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        String sheetName = "Car";
        Workbook workbook = WorkbookUtility.open(excelFile);
        Assertions.assertEquals(true, SheetUtility.isNull(workbook, sheetName));
    }

    @Test
    void testIsNull() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        Workbook workbook = WorkbookUtility.open(excelFile);
        Assertions.assertEquals(false, SheetUtility.isNull(workbook, 1));
    }
}