package io.github.mbenincasa.javaexcelutils.model.excel;

import io.github.mbenincasa.javaexcelutils.exceptions.CellNotFoundException;
import io.github.mbenincasa.javaexcelutils.exceptions.RowNotFoundException;
import io.github.mbenincasa.javaexcelutils.tools.ExcelUtility;
import lombok.AllArgsConstructor;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.LinkedList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

/**
 * {@code ExcelSheet} is the {@code Sheet} wrapper class of the Apache POI library
 * @author Mirko Benincasa
 * @since 0.3.0
 */
@AllArgsConstructor
@Getter
@EqualsAndHashCode
public class ExcelSheet {

    /**
     * The Sheet index in the Workbook
     * @return The Sheet index
     */
    @SneakyThrows
    public Integer getIndex() {
        if (this.sheet == null)
            return null;
        return getWorkbook().getWorkbook().getSheetIndex(this.name);
    }

    /**
     * This object refers to the Apache POI Library {@code Sheet}
     */
    private Sheet sheet;

    /**
     * The index of the Sheet in the Workbook
     */
    private Integer index;

    /**
     * The name of the Sheet
     */
    private String name;

    /**
     * Returns the Workbook to which it belongs
     * @return A ExcelWorkbook
     */
    public ExcelWorkbook getWorkbook() {
        return new ExcelWorkbook(this.getSheet().getWorkbook());
    }

    /**
     * Remove the selected Sheet
     * @since 0.4.1
     */
    public void remove() {
        getWorkbook().removeSheet(getIndex());
        this.sheet = null;
        this.name = null;
        this.index = null;
    }

    /**
     * The list of Rows related to the Sheet
     * @return A list of Rows
     */
    public List<ExcelRow> getRows() {
        List<ExcelRow> excelRows = new LinkedList<>();
        for (Row row : this.sheet) {
            excelRows.add(new ExcelRow(row, row.getRowNum()));
        }

        return excelRows;
    }

    /**
     * Retrieve a row by index
     * @param index The index of the row requested
     * @return A ExcelRow
     * @throws RowNotFoundException If the row is not present or has not been created
     * @since 0.4.1
     */
    public ExcelRow getRow(Integer index) throws RowNotFoundException {
        Row row = this.sheet.getRow(index);
        if (row == null) {
            throw new RowNotFoundException("There is not a row in the index: " + index);
        }
        return new ExcelRow(row, index);
    }

    /**
     * Retrieve or create a row by index
     * @param index The index of the row requested
     * @return A ExcelRow
     * @since 0.4.1
     */
    public ExcelRow getOrCreateRow(Integer index) {
        Row row = this.sheet.getRow(index);
        if (row == null) {
            return createRow(index);
        }
        return new ExcelRow(row, index);
    }

    /**
     * Removes a row by index
     * @param index The index of the row to remove
     * @throws RowNotFoundException If the row is not present or has not been created
     * @since 0.4.1
     */
    public void removeRow(Integer index) throws RowNotFoundException {
        ExcelRow excelRow = getRow(index);
        this.sheet.removeRow(excelRow.getRow());
    }

    /**
     * Method that allows you to write a data matrix starting from a source cell
     * @param startingCell Source cell
     * @param data A matrix data
     * @since 0.5.0
     */
    public void writeCells(String startingCell, Stream<Object[]> data) {
        int[] cellIndexes = ExcelUtility.getCellIndexes(startingCell);
        AtomicInteger rowIndex = new AtomicInteger(cellIndexes[0]);
        int colIndex = cellIndexes[1];
        writeOrAppendCells(data, rowIndex, colIndex);
    }

    /**
     * Method that allows you to append a data matrix starting from a source cell
     * @param startingCell Source cell
     * @param data A matrix data
     * @since 0.5.0
     */
    public void appendCells(String startingCell, Stream<Object[]> data) {
        int[] cellIndexes = ExcelUtility.getCellIndexes(startingCell);
        AtomicInteger rowIndex = new AtomicInteger(getLastRowIndex() + 1);
        int colIndex = cellIndexes[1];
        writeOrAppendCells(data, rowIndex, colIndex);
    }

    private void writeOrAppendCells(Stream<Object[]> data, AtomicInteger rowIndex, int colIndex) {
        data.forEach(rowData -> {
            ExcelRow excelRow = getOrCreateRow(rowIndex.get());
            for (int i = 0; i < rowData.length; i++) {
                Object value = rowData[i];
                ExcelCell excelCell = excelRow.getOrCreateCell(colIndex + i);
                excelCell.writeValue(value);
            }
            rowIndex.getAndIncrement();
        });
    }

    /**
     * Method used to delete cells based on the selected range, for example: 'A1:B2'
     * @param cellRange The range of cells to be deleted
     * @since 0.5.0
     */
    public void removeCells(String cellRange) {
        Pattern pattern = Pattern.compile("(\\w+):(\\w+)");
        Matcher matcher = pattern.matcher(cellRange);
        if (!matcher.matches())
            throw new IllegalArgumentException("Invalid input format: " + cellRange + ". Provide input for example: 'A1:B2'");
        CellRangeAddress range = CellRangeAddress.valueOf(cellRange);
        for (int rowIndex = range.getFirstRow(); rowIndex <= range.getLastRow(); rowIndex++) {
            try {
                ExcelRow excelRow = getRow(rowIndex);
                for (int colIndex = range.getFirstColumn(); colIndex <= range.getLastColumn(); colIndex++) {
                    ExcelCell excelCell = excelRow.getCell(colIndex);
                    excelCell.remove();
                }
            } catch (RowNotFoundException | CellNotFoundException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * Create a new Row in the Sheet
     * @param index The index in the Sheet
     * @return A Row
     */
    public ExcelRow createRow(Integer index) {
        return new ExcelRow(this.sheet.createRow(index), index);
    }

    /**
     * Retrieves the index of the last Row
     * @return The index of the last Row
     */
    public Integer getLastRowIndex() {
        return this.sheet.getLastRowNum();
    }

    /**
     * Counts how many Rows are compiled
     * @param alsoEmpty {@code true} if you want to count Rows that have all empty Cells
     * @return The number of Rows compiled
     */
    public Integer countAllRows(Boolean alsoEmpty) {
        Integer count = this.getLastRowIndex() + 1;
        if (alsoEmpty)
            return count;

        for (int i = 0; i < this.sheet.getPhysicalNumberOfRows(); i++) {
            Row row = this.sheet.getRow(i);
            boolean isEmptyRow = true;

            if (row == null) {
                count--;
                continue;
            }

            for (int j = 0; j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                if (cell != null) {
                    Object val;
                    switch (cell.getCellType()) {
                        case NUMERIC -> val = cell.getNumericCellValue();
                        case BOOLEAN -> val = cell.getBooleanCellValue();
                        default -> val = cell.getStringCellValue();
                    }
                    if (val != null) {
                        isEmptyRow = false;
                        break;
                    }
                }
            }

            if (isEmptyRow) {
                count--;
            }
        }

        return count;
    }
}
