package io.github.mbenincasa.javaexcelutils.model;

import io.github.mbenincasa.javaexcelutils.exceptions.ExtensionNotValidException;
import io.github.mbenincasa.javaexcelutils.exceptions.OpenWorkbookException;
import io.github.mbenincasa.javaexcelutils.exceptions.SheetNotFoundException;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;

class ExcelRowTest {

    private final File excelFile = new File("./src/test/resources/employee.xlsx");

    @Test
    void getCells() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet("Office");
        ExcelRow excelRow = excelSheet.getRows().get(0);
        List<ExcelCell> excelCells = excelRow.getCells();
        Assertions.assertEquals("CITY", excelCells.get(0).getCell().getStringCellValue());
        Assertions.assertEquals("PROVINCE", excelCells.get(1).getCell().getStringCellValue());
        Assertions.assertEquals("NUMBER OF STATIONS", excelCells.get(2).getCell().getStringCellValue());

    }

    @Test
    void getSheet() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet();
        ExcelRow excelRow = excelSheet.getRows().get(0);
        ExcelSheet excelSheet1 = excelRow.getSheet();
        Assertions.assertEquals(excelSheet, excelSheet1);
    }

    @Test
    void createCell() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheetOrCreate("TestWrite");
        ExcelRow excelRow = excelSheet.createRow(0);
        ExcelCell excelCell = excelRow.createCell(0);
        Assertions.assertNotNull(excelCell.getCell());
        Assertions.assertEquals(0, excelCell.getIndex());
    }

    @Test
    void getLastColumnIndex() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(1);
        List<ExcelRow> excelRows = excelSheet.getRows();
        Assertions.assertEquals(4, excelRows.get(0).getLastColumnIndex());
        Assertions.assertEquals(2, excelRows.get(1).getLastColumnIndex());
    }

    @Test
    void countAllColumns() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(1);
        List<ExcelRow> excelRows = excelSheet.getRows();
        Assertions.assertEquals(4, excelRows.get(0).countAllColumns(false));
        Assertions.assertEquals(3, excelRows.get(1).countAllColumns(true));
    }
}