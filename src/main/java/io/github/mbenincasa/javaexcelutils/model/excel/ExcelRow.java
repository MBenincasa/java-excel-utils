package io.github.mbenincasa.javaexcelutils.model.excel;

import io.github.mbenincasa.javaexcelutils.exceptions.CellNotFoundException;
import io.github.mbenincasa.javaexcelutils.exceptions.ReadValueException;
import io.github.mbenincasa.javaexcelutils.exceptions.RowNotFoundException;
import lombok.*;
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
@AllArgsConstructor(access = AccessLevel.PRIVATE)
@Getter
@EqualsAndHashCode
@Builder(access = AccessLevel.PRIVATE)
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
     * Get to an ExcelRow instance from Apache POI Row
     * @param row The Row instance to wrap
     * @return The ExcelRow instance
     * @since 0.5.0
     */
    public static ExcelRow of(Row row) {
        return ExcelRow.builder()
                .row(row)
                .index(row.getRowNum())
                .build();
    }

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
            excelCells.add(ExcelCell.of(cell));
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
        return ExcelCell.of(cell);
    }

    /**
     * Retrieve or create a cell by index
     * @param index The index of the cell requested
     * @return A ExcelCell
     * @since 0.4.1
     */
    public ExcelCell getOrCreateCell(Integer index) {
        Cell cell = this.row.getCell(index);
        if (cell == null) {
            return  createCell(index);
        }
        return ExcelCell.of(cell);
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
     * Write the values in the cells of the row
     * @param values The values to write in the cells of the row
     * @since 0.4.1
     */
    public void writeValues(List<?> values) {
        for (int i = 0; i < values.size(); i++) {
            ExcelCell excelCell = getOrCreateCell(i);
            excelCell.writeValue(values.get(i));
        }
    }

    /**
     * Reads the values of all cells in the row
     * @return The list of values written in the cells
     * @throws ReadValueException If an error occurs while reading
     * @since 0.4.1
     */
    public List<?> readValues() throws ReadValueException {
        List<Object> values = new LinkedList<>();
        for (ExcelCell excelCell : getCells()) {
            values.add(excelCell.readValue());
        }
        return values;
    }

    /**
     * Reads the values of all cells in the row
     * @param classes A list of Classes that is used to cast the results read
     * @return The list of values written in the cells
     * @throws ReadValueException If an error occurs while reading
     * @since 0.4.1
     */
    public List<?> readValues(List<Class<?>> classes) throws ReadValueException {
        List<ExcelCell> excelCells = getCells();
        if(excelCells.size() != classes.size()) {
            throw new IllegalArgumentException("There are " + excelCells.size() + " items in the row and classlist has " + classes.size() + " values.");
        }

        List<Object> values = new LinkedList<>();
        for (int i = 0; i < excelCells.size(); i++) {
            values.add(excelCells.get(i).readValue(classes.get(i)));
        }

        return values;
    }

    /**
     * Reads the values of all cells in the row as a String
     * @return The list of values, such as String, written in the cells
     * @since 0.4.1
     */
    public List<String> readValuesAsString() {
        List<String> values = new LinkedList<>();
        for (ExcelCell excelCell : getCells()) {
            values.add(excelCell.readValueAsString());
        }
        return values;
    }

    /**
     * Returns the Sheet to which it belongs
     * @return A ExcelSheet
     */
    @SneakyThrows
    public ExcelSheet getSheet() {
        Sheet sheet = this.row.getSheet();
        return ExcelSheet.of(sheet);
    }

    /**
     * Create a new Cell in the Row
     * @param index The index in the Row
     * @return A Cell
     */
    public ExcelCell createCell(Integer index) {
        return ExcelCell.of(this.row.createCell(index));
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
