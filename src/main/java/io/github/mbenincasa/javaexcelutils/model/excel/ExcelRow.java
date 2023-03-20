package io.github.mbenincasa.javaexcelutils.model.excel;

import io.github.mbenincasa.javaexcelutils.exceptions.CellNotFoundException;
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
 * {@code ExcelRow} is the {@code Row} wrapper class of the Apache POI library
 * @author Mirko Benincasa
 * @since 0.3.0
 */
@AllArgsConstructor
@Getter
@EqualsAndHashCode
public class ExcelRow {

    /**
     * This object refers to the Apache POI Library {@code Row}
     */
    private Row row;

    /**
     * The index of the Row in the Sheet
     */
    private Integer index;

    /**
     * Remove the selected Row
     * @throws RowNotFoundException If the row is not present or has not been created
     * @since 0.4.1
     */
    public void remove() throws RowNotFoundException {
        getSheet().removeRow(this.index);
        this.row = null;
        this.index = null;
    }

    /**
     * The list of Cells related to the Row
     * @return A list of Cells
     */
    public List<ExcelCell> getCells() {
        List<ExcelCell> excelCells = new LinkedList<>();
        for (Cell cell : this.row) {
            excelCells.add(new ExcelCell(cell, cell.getColumnIndex()));
        }
        return excelCells;
    }

    /**
     * Retrieve a cell by index
     * @param index The index of the cell requested
     * @return A ExcelCell
     * @throws CellNotFoundException If the cell is not present or has not been created
     * @since 0.4.1
     */
    public ExcelCell getCell(Integer index) throws CellNotFoundException {
        Cell cell = this.row.getCell(index);
        if (cell == null) {
            throw new CellNotFoundException("There is not a cell in the index: " + index);
        }
        return new ExcelCell(cell, cell.getColumnIndex());
    }

    /**
     * Removes a cell by index
     * @param index The index of the row to remove
     * @throws CellNotFoundException If the cell is not present or has not been created
     * @since 0.4.1
     */
    public void removeCell(Integer index) throws CellNotFoundException {
        ExcelCell excelCell = getCell(index);
        this.row.removeCell(excelCell.getCell());
    }

    /**
     * Returns the Sheet to which it belongs
     * @return A ExcelSheet
     */
    @SneakyThrows
    public ExcelSheet getSheet() {
        Sheet sheet = this.row.getSheet();
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(sheet.getWorkbook());
        String sheetName = sheet.getSheetName();
        return new ExcelSheet(sheet, excelWorkbook.getSheet(sheetName).getIndex(), sheetName);
    }

    /**
     * Create a new Cell in the Row
     * @param index The index in the Row
     * @return A Cell
     */
    public ExcelCell createCell(Integer index) {
        return new ExcelCell(this.row.createCell(index), index);
    }

    /**
     * Retrieves the index of the last Cell
     * @return The index of the last Cell
     */
    public Integer getLastColumnIndex() {
        return this.row.getLastCellNum() - 1;
    }

    /**
     * Counts how many Cells are compiled
     * @param alsoEmpty {@code true} if you want to count Cells empty
     * @return The number of Cells compiled
     */
    public Integer countAllColumns(Boolean alsoEmpty) {
        Integer count = this.getLastColumnIndex() + 1;
        if (alsoEmpty)
            return count;

        for (int i = 0; i < this.row.getPhysicalNumberOfCells(); i++) {
            Cell cell = this.row.getCell(i);
            if (cell == null) {
                count--;
                continue;
            }

            Object val;
            switch (cell.getCellType()) {
                case NUMERIC -> val = cell.getNumericCellValue();
                case BOOLEAN -> val = cell.getBooleanCellValue();
                default -> val = cell.getStringCellValue();
            }

            if (val == null) {
                count--;
            }
        }

        return count;
    }
}
