package io.github.mbenincasa.javaexcelutils.model.excel;

import io.github.mbenincasa.javaexcelutils.exceptions.*;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.List;

class ExcelCellTest {

    private final File excelFile = new File("./src/test/resources/employee.xlsx");

    @Test
    void getRow() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet();
        ExcelRow excelRow = excelSheet.getRows().get(0);
        ExcelCell excelCell = excelRow.getCells().get(0);
        ExcelRow excelRow1 = excelCell.getRow();
        Assertions.assertEquals(excelRow, excelRow1);
    }

    @Test
    void readValue() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, ReadValueException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet();
        ExcelRow excelRow = excelSheet.getRows().get(1);
        List<ExcelCell> excelCells = excelRow.getCells();
        Assertions.assertEquals("Rossi", excelCells.get(0).readValue(String.class));
        Assertions.assertEquals("Mario", excelCells.get(1).readValue(String.class));
        Assertions.assertEquals(25, excelCells.get(2).readValue(Integer.class));
        Assertions.assertEquals(LocalDate.of(1987, 5, 22), excelCells.get(3).readValue(LocalDate.class));
        Assertions.assertNotNull(excelCells.get(4).readValue(Date.class));
        Assertions.assertEquals(28000.00, excelCells.get(5).readValue(Double.class));
        Assertions.assertEquals(LocalDateTime.of(2023, 2, 11, 12, 35, 55, 603000000), excelCells.get(6).readValue(LocalDateTime.class));
        Assertions.assertEquals(true, excelCells.get(7).readValue(Boolean.class));
    }

    @Test
    void writeValue() throws OpenWorkbookException, ExtensionNotValidException, IOException, ReadValueException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheetOrCreate("TestWrite");
        ExcelRow excelRow = excelSheet.createRow(0);
        ExcelCell excelCell = excelRow.createCell(0);
        excelCell.writeValue("Text");
        ExcelCell excelCell1 = excelRow.createCell(1);
        excelCell1.writeValue(21);
        ExcelCell excelCell2 = excelRow.createCell(2);
        LocalDateTime localDateTime = LocalDateTime.of(2021, 1, 1, 21, 21, 21, 0);
        excelCell2.writeValue(localDateTime);
        ExcelCell excelCell3 = excelRow.createCell(3);
        excelCell3.writeValue(false);
        Assertions.assertEquals("Text", excelCell.readValue(String.class));
        Assertions.assertEquals(21, excelCell1.readValue(Integer.class));
        Assertions.assertEquals(localDateTime, excelCell2.readValue(LocalDateTime.class));
        Assertions.assertEquals(false, excelCell3.readValue(Boolean.class));
    }

    @Test
    void formatStyle() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet();
        ExcelRow excelRow = excelSheet.getRows().get(0);
        ExcelCell excelCell = excelRow.getCells().get(0);
        excelCell.formatStyle((short) 1);
        Assertions.assertEquals((short) 1, excelCell.getCell().getCellStyle().getDataFormat());
    }

    @Test
    void testReadValue() throws OpenWorkbookException, ExtensionNotValidException, IOException, ReadValueException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheetOrCreate("TestWrite");
        ExcelRow excelRow = excelSheet.createRow(0);
        ExcelCell excelCell = excelRow.createCell(0);
        excelCell.writeValue("Text");
        ExcelCell excelCell1 = excelRow.createCell(1);
        excelCell1.writeValue(21);
        ExcelCell excelCell2 = excelRow.createCell(2);
        LocalDateTime localDateTime = LocalDateTime.of(2021, 1, 1, 21, 21, 21, 0);
        excelCell2.writeValue(localDateTime);
        ExcelCell excelCell3 = excelRow.createCell(3);
        excelCell3.writeValue(false);
        Assertions.assertEquals("Text", excelCell.readValue());
        Assertions.assertEquals(21.0, excelCell1.readValue());
        Assertions.assertEquals(44197.88982638889, excelCell2.readValue());
        Assertions.assertEquals(false, excelCell3.readValue());
    }

    @Test
    void readValueAsString() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheetOrCreate("TestWrite");
        ExcelRow excelRow = excelSheet.createRow(0);
        ExcelCell excelCell = excelRow.createCell(0);
        excelCell.writeValue("Text");
        ExcelCell excelCell1 = excelRow.createCell(1);
        excelCell1.writeValue(21);
        ExcelCell excelCell2 = excelRow.createCell(2);
        LocalDateTime localDateTime = LocalDateTime.of(2021, 1, 1, 21, 21, 21, 0);
        excelCell2.writeValue(localDateTime);
        ExcelCell excelCell3 = excelRow.createCell(3);
        excelCell3.writeValue(false);
        Assertions.assertEquals("Text", excelCell.readValueAsString());
        Assertions.assertEquals("21", excelCell1.readValueAsString());
        Assertions.assertEquals("1/1/21 21:21", excelCell2.readValueAsString());
        Assertions.assertEquals("FALSE", excelCell3.readValueAsString());
    }

    @Test
    void remove() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, RowNotFoundException, CellNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet();
        ExcelRow excelRow = excelSheet.getRow(0);
        ExcelCell excelCell = excelRow.getCell(0);
        Assertions.assertDoesNotThrow(excelCell::remove);
        Assertions.assertNull(excelCell.getCell());
        Assertions.assertNull(excelCell.getIndex());
    }

    @Test
    void getCellName() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, RowNotFoundException, CellNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet();
        ExcelRow excelRow = excelSheet.getRow(0);
        ExcelCell excelCell = excelRow.getCell(0);
        String cellName = excelCell.getCellName();
        Assertions.assertEquals("A1", cellName);
    }
}