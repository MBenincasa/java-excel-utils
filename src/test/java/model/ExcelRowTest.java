package model;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;
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
}