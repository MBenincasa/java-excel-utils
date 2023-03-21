package io.github.mbenincasa.javaexcelutils.model.excel;

import io.github.mbenincasa.javaexcelutils.exceptions.*;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedList;
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

    @Test
    void remove() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, RowNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelRow excelRow = excelSheet.getRow(0);
        Assertions.assertDoesNotThrow(excelRow::remove);
        Assertions.assertNull(excelRow.getRow());
        Assertions.assertNull(excelRow.getIndex());
    }

    @Test
    void getCell() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, RowNotFoundException, CellNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelRow excelRow = excelSheet.getRow(0);
        ExcelCell excelCell = excelRow.getCell(0);
        Assertions.assertEquals(0, excelCell.getIndex());
        Assertions.assertNotNull(excelCell.getCell());
    }

    @Test
    void removeCell() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, RowNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelRow excelRow = excelSheet.getRow(0);
        Assertions.assertDoesNotThrow(() -> excelRow.removeCell(0));
        Assertions.assertThrows(CellNotFoundException.class, () -> excelRow.getCell(0));
    }

    @Test
    void getOrCreateCell() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, RowNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelRow excelRow = excelSheet.getRow(0);
        ExcelCell excelCell = excelRow.getOrCreateCell(20);
        Assertions.assertEquals(20, excelCell.getIndex());
        Assertions.assertNotNull(excelCell.getCell());
    }

    @Test
    void writeValues() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, RowNotFoundException, ReadValueException {
        List<Object> values = new ArrayList<>();
        values.add("Rossi");
        values.add(3);
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(0);
        ExcelRow excelRow = excelSheet.getRow(0);
        excelRow.writeValues(values);
        List<ExcelCell> excelCells = excelRow.getCells();
        Assertions.assertEquals("Rossi", excelCells.get(0).readValue(String.class));
        Assertions.assertEquals(3, excelCells.get(1).readValue(Integer.class));
    }

    @Test
    void readValues() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, RowNotFoundException, ReadValueException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(1);
        ExcelRow excelRow = excelSheet.getRow(1);
        List<?> values = excelRow.readValues();
        Assertions.assertEquals("Nocera Inferiore", values.get(0));
        Assertions.assertEquals("Salerno", values.get(1));
        Assertions.assertEquals(40.0, values.get(2));
    }

    @Test
    void testReadValues() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, RowNotFoundException, ReadValueException {
        List<Class<?>> classes = new LinkedList<>();
        classes.add(String.class);
        classes.add(String.class);
        classes.add(Integer.class);
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(1);
        ExcelRow excelRow = excelSheet.getRow(1);
        List<?> values = excelRow.readValues(classes);
        Assertions.assertEquals("Nocera Inferiore", values.get(0));
        Assertions.assertEquals("Salerno", values.get(1));
        Assertions.assertEquals(40, values.get(2));
    }

    @Test
    void readValuesAsString() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, RowNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet(1);
        ExcelRow excelRow = excelSheet.getRow(1);
        List<?> values = excelRow.readValuesAsString();
        Assertions.assertEquals("Nocera Inferiore", values.get(0));
        Assertions.assertEquals("Salerno", values.get(1));
        Assertions.assertEquals("40", values.get(2));
    }
}