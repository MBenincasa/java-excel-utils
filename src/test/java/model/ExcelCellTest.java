package model;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;

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
}