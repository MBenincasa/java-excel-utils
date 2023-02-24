package model;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetAlreadyExistsException;
import exceptions.SheetNotFoundException;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;

public class ExcelSheetTest {

    private final File excelFile = new File("./src/test/resources/employee.xlsx");

    @Test
    void create() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetAlreadyExistsException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = ExcelSheet.create(excelWorkbook);
        Assertions.assertNotNull(excelSheet.getSheet());
    }

    @Test
    void testCreate() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetAlreadyExistsException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        String sheetName = "Admin";
        ExcelSheet excelSheet = ExcelSheet.create(excelWorkbook, sheetName);
        Assertions.assertNotNull(excelSheet.getSheet());
        Assertions.assertEquals(true, excelWorkbook.isSheetPresent(sheetName));
    }

    @Test
    void getWorkbook() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelWorkbook excelWorkbook1 = excelSheet.getWorkbook();
        Assertions.assertNotNull(excelWorkbook1.getWorkbook());
    }

    @Test
    void getRows() {
    }
}