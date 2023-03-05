package io.github.mbenincasa.javaexcelutils.model;

import io.github.mbenincasa.javaexcelutils.exceptions.ExtensionNotValidException;
import io.github.mbenincasa.javaexcelutils.exceptions.OpenWorkbookException;
import io.github.mbenincasa.javaexcelutils.exceptions.SheetNotFoundException;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelRow;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelSheet;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelWorkbook;
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
}