package io.github.mbenincasa.javaexcelutils.model.excel;

import io.github.mbenincasa.javaexcelutils.exceptions.ExtensionNotValidException;
import io.github.mbenincasa.javaexcelutils.exceptions.OpenWorkbookException;
import io.github.mbenincasa.javaexcelutils.exceptions.RowNotFoundException;
import io.github.mbenincasa.javaexcelutils.exceptions.SheetNotFoundException;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class ExcelSheetTest {

    private final File excelFile = new File("./src/test/resources/employee.xlsx");

    @Test
    void getWorkbook() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelWorkbook excelWorkbook1 = excelSheet.getWorkbook();
        Assertions.assertNotNull(excelWorkbook1.getWorkbook());
    }

    @Test
    void getRows() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        List<ExcelRow> excelRows = excelSheet.getRows();
        Assertions.assertEquals(3, excelRows.size());
    }

    @Test
    void createRow() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheetOrCreate("TestWrite");
        ExcelRow excelRow = excelSheet.createRow(0);
        Assertions.assertNotNull(excelRow.getRow());
        Assertions.assertEquals(0, excelRow.getIndex());
    }

    @Test
    void getLastRowIndex() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        Assertions.assertEquals(2, excelSheet.getLastRowIndex());
    }

    @Test
    void countAllRows() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(1);
        Assertions.assertEquals(4, excelSheet.countAllRows(false));
    }

    @Test
    void getIndex() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet("Employee");
        ExcelSheet excelSheet1 = excelWorkbook.getSheet("Office");
        Assertions.assertEquals(0, excelSheet.getIndex());
        Assertions.assertEquals(1, excelSheet1.getIndex());
    }

    @Test
    void remove() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelSheet excelSheet1 = excelWorkbook.getSheet(1);
        Assertions.assertDoesNotThrow(excelSheet::remove);
        Assertions.assertEquals(0, excelSheet1.getIndex());
        Assertions.assertNull(excelSheet.getIndex());
        Assertions.assertNull(excelSheet.getName());
        Assertions.assertNull(excelSheet.getSheet());
    }

    @Test
    void getRow() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, RowNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelRow excelRow = excelSheet.getRow(0);
        Assertions.assertEquals(0, excelRow.getIndex());
        Assertions.assertNotNull(excelRow.getRow());
    }

    @Test
    void removeRow() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        Assertions.assertDoesNotThrow(() -> excelSheet.removeRow(0));
        Assertions.assertThrows(RowNotFoundException.class, () -> excelSheet.getRow(0));
    }

    @Test
    void getOrCreateRow() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelRow excelRow = excelSheet.getOrCreateRow(20);
        Assertions.assertEquals(20, excelRow.getIndex());
        Assertions.assertNotNull(excelRow.getRow());
    }
}