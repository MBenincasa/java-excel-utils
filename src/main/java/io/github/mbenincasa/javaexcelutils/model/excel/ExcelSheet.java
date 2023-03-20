package io.github.mbenincasa.javaexcelutils.model.excel;

import io.github.mbenincasa.javaexcelutils.exceptions.RowNotFoundException;
import lombok.AllArgsConstructor;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.LinkedList;
import java.util.List;

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
