package io.github.mbenincasa.javaexcelutils.model.excel;

import io.github.mbenincasa.javaexcelutils.annotations.ExcelCellMapping;
import io.github.mbenincasa.javaexcelutils.exceptions.CellNotFoundException;
import io.github.mbenincasa.javaexcelutils.exceptions.ReadValueException;
import io.github.mbenincasa.javaexcelutils.exceptions.RowNotFoundException;
import io.github.mbenincasa.javaexcelutils.model.parser.Direction;
import io.github.mbenincasa.javaexcelutils.model.parser.ExcelCellParser;
import io.github.mbenincasa.javaexcelutils.model.parser.ExcelListParserMapping;
import io.github.mbenincasa.javaexcelutils.tools.ExcelUtility;
import lombok.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
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
@AllArgsConstructor(access = AccessLevel.PRIVATE)
@Getter
@EqualsAndHashCode
@Builder(access = AccessLevel.PRIVATE)
public class ExcelSheet {

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
     * Get to an ExcelSheet instance from Apache POI Sheet
     * @param sheet The Sheet instance to wrap
     * @return The ExcelSheet instance
     * @since 0.5.0
     */
    public static ExcelSheet of(Sheet sheet) {
        return ExcelSheet.builder()
                .sheet(sheet)
                .index(sheet.getWorkbook().getSheetIndex(sheet))
                .name(sheet.getSheetName())
                .build();
    }

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
     * Returns the Workbook to which it belongs
     * @return A ExcelWorkbook
     */
    public ExcelWorkbook getWorkbook() {
        return ExcelWorkbook.of(this.getSheet().getWorkbook());
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
            excelRows.add(ExcelRow.of(row));
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
        return ExcelRow.of(row);
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
        return ExcelRow.of(row);
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
        return ExcelRow.of(this.sheet.createRow(index));
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

    /**
     * @param clazz The class of the object to return
     * @param startingCell The name of the source cell
     * @param <T> The class parameter of the object
     * @return The object parsed
     * @throws NoSuchMethodException If the setting method or empty constructor of the object is not found
     * @throws InvocationTargetException If an error occurs while instantiating a new object or setting a field
     * @throws InstantiationException If an error occurs while instantiating a new object
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws ReadValueException If an error occurs while reading a cell
     * @since 0.5.0
     */
    public <T> T parseToObject(Class<T> clazz, String startingCell) throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException, ReadValueException {
        T obj = clazz.getDeclaredConstructor().newInstance();
        int[] cellIndexes = ExcelUtility.getCellIndexes(startingCell);
        int startingRow = cellIndexes[0];
        int startingCol = cellIndexes[1];
        Field[] fields = clazz.getDeclaredFields();

        for (Field field : fields) {
            if (!field.isAnnotationPresent(ExcelCellMapping.class))
                continue;

            ExcelCellMapping excelCellMapping = field.getAnnotation(ExcelCellMapping.class);
            int deltaRow = excelCellMapping.deltaRow();
            int deltaCol = excelCellMapping.deltaCol();
            ExcelRow excelRow = getOrCreateRow(startingRow + deltaRow);
            ExcelCell excelCell = excelRow.getOrCreateCell(startingCol + deltaCol);
            field.setAccessible(true);

            Object value = ExcelCellParser.class.isAssignableFrom(field.getType())
                    ? parseToObject(field.getType(), ExcelUtility.getCellName(startingRow + deltaRow, startingCol + deltaCol))
                    : excelCell.readValue(field.getType());
            field.set(obj, value);

        }

        return obj;
    }

    /**
     * @param clazz The class of the object to return
     * @param mapping The rules to retrieve the list of objects
     * @param <T> The class parameter of the object
     * @return The object list parsed
     * @throws ReadValueException If an error occurs while reading a cell
     * @throws InvocationTargetException If an error occurs while instantiating a new object or setting a field
     * @throws NoSuchMethodException If the setting method or empty constructor of the object is not found
     * @throws InstantiationException If an error occurs while instantiating a new object
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @since 0.5.0
     */
    public <T> List<T> parseToList(Class<T> clazz, ExcelListParserMapping mapping) throws ReadValueException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        List<T> objectList = new LinkedList<>();
        int[] cellIndexes = ExcelUtility.getCellIndexes(mapping.getStartingCell());
        int startingRow = cellIndexes[0];
        int startingCol = cellIndexes[1];

        if (mapping.getDirection() == Direction.HORIZONTAL) {
            int maxCol = getOrCreateRow(startingRow).getLastColumnIndex();
            while (startingCol <= maxCol) {
                String currentCell = ExcelUtility.getCellName(startingRow, startingCol);
                T obj = parseToObject(clazz, currentCell);
                objectList.add(obj);
                startingCol += mapping.getJumpCells();
            }
        } else if (mapping.getDirection() == Direction.VERTICAL) {
            int maxRow = getLastRowIndex();
            while (startingRow <= maxRow) {
                String currentCell = ExcelUtility.getCellName(startingRow, startingCol);
                T obj = parseToObject(clazz, currentCell);
                objectList.add(obj);
                startingRow += mapping.getJumpCells();
            }
        }

        return objectList;
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
}
