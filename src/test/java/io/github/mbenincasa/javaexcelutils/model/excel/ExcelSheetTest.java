package io.github.mbenincasa.javaexcelutils.model.excel;

import io.github.mbenincasa.javaexcelutils.exceptions.*;
import io.github.mbenincasa.javaexcelutils.tools.utils.ParsableEmployee;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.time.LocalDate;
import java.util.List;
import java.util.stream.Stream;

public class ExcelSheetTest {

    private final File excelFile = new File("./src/test/resources/employee.xlsx");
    private final File excelFileToParse = new File("./src/test/resources/parse_to_object.xlsx");

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

    @Test
    void removeCells() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet("Test_2");
        excelSheet.removeCells("B2:C3");
        List<ExcelRow> excelRows = excelSheet.getRows();
        Assertions.assertThrows(CellNotFoundException.class, () -> excelRows.get(1).getCell(1));
        Assertions.assertThrows(CellNotFoundException.class, () -> excelRows.get(1).getCell(2));
        Assertions.assertThrows(CellNotFoundException.class, () -> excelRows.get(2).getCell(1));
        Assertions.assertThrows(CellNotFoundException.class, () -> excelRows.get(2).getCell(2));
    }

    @Test
    void writeCells() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, RowNotFoundException, CellNotFoundException, ReadValueException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet("Test_2");

        Object[] row1 = {"Nome", "Cognome", "Età"};
        Object[] row2 = {"Mario", "Rossi", 30};
        Stream<Object[]> rows = Stream.of(row1, row2);
        excelSheet.writeCells("C5", rows);
        ExcelRow excelRow5 = excelSheet.getRow(4);
        Assertions.assertEquals("Nome", excelRow5.getCell(2).readValue(String.class));
        Assertions.assertEquals("Cognome", excelRow5.getCell(3).readValue(String.class));
        Assertions.assertEquals("Età", excelRow5.getCell(4).readValue(String.class));
        ExcelRow excelRow6 = excelSheet.getRow(5);
        Assertions.assertEquals("Mario", excelRow6.getCell(2).readValue(String.class));
        Assertions.assertEquals("Rossi", excelRow6.getCell(3).readValue(String.class));
        Assertions.assertEquals(30, excelRow6.getCell(4).readValue(Integer.class));
    }

    @Test
    void appendCells() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, RowNotFoundException, CellNotFoundException, ReadValueException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        ExcelSheet excelSheet = excelWorkbook.getSheet("Test_2");

        Object[] row1 = {"Nome", "Cognome", "Età"};
        Object[] row2 = {"Mario", "Rossi", 30};
        Stream<Object[]> rows = Stream.of(row1, row2);
        excelSheet.appendCells("B1", rows);
        ExcelRow excelRow5 = excelSheet.getRow(4);
        Assertions.assertEquals("Nome", excelRow5.getCell(1).readValue(String.class));
        Assertions.assertEquals("Cognome", excelRow5.getCell(2).readValue(String.class));
        Assertions.assertEquals("Età", excelRow5.getCell(3).readValue(String.class));
        ExcelRow excelRow6 = excelSheet.getRow(5);
        Assertions.assertEquals("Mario", excelRow6.getCell(1).readValue(String.class));
        Assertions.assertEquals("Rossi", excelRow6.getCell(2).readValue(String.class));
        Assertions.assertEquals(30, excelRow6.getCell(3).readValue(Integer.class));
    }

    @Test
    void of() throws OpenWorkbookException, ExtensionNotValidException, IOException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFile);
        Workbook workbook = excelWorkbook.getWorkbook();
        Sheet sheet = workbook.getSheet("Test_2");
        ExcelSheet excelSheet = ExcelSheet.of(sheet);
        Assertions.assertEquals(sheet, excelSheet.getSheet());
        Assertions.assertEquals(sheet.getSheetName(), excelSheet.getName());
        Assertions.assertEquals(sheet.getWorkbook().getSheetIndex(sheet), excelSheet.getIndex());
    }

    @Test
    void parseToObject() throws OpenWorkbookException, ExtensionNotValidException, IOException, SheetNotFoundException, ReadValueException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(excelFileToParse);
        ExcelSheet excelSheet = excelWorkbook.getSheet("DATA");
        ParsableEmployee employee = excelSheet.parseToObject(ParsableEmployee.class, "A1");
        Assertions.assertEquals("Mario", employee.getName());
        Assertions.assertEquals("Rossi", employee.getLastName());
        Assertions.assertEquals(25, employee.getAge());
        Assertions.assertEquals(LocalDate.of(2022, 1 , 12), employee.getHireDate());
        Assertions.assertEquals(LocalDate.of(2022, 2 , 12), employee.getTerminationDate());
        Assertions.assertEquals("Nocera Inferiore", employee.getAddress().getCity());
        Assertions.assertEquals("84014", employee.getAddress().getCap());
    }
}