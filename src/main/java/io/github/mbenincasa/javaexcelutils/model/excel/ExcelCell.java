package io.github.mbenincasa.javaexcelutils.model.excel;

import io.github.mbenincasa.javaexcelutils.exceptions.CellNotFoundException;
import io.github.mbenincasa.javaexcelutils.exceptions.ReadValueException;
import io.github.mbenincasa.javaexcelutils.tools.ExcelUtility;
import lombok.*;
import org.apache.poi.ss.usermodel.*;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

/**
 * {@code ExcelCell} is the {@code Cell} wrapper class of the Apache POI library
 * @author Mirko Benincasa
 * @since 0.3.0
 */
@AllArgsConstructor(access = AccessLevel.PRIVATE)
@Getter
@EqualsAndHashCode
@Builder(access = AccessLevel.PRIVATE)
public class ExcelCell {

    /**
     * This object refers to the Apache POI Library {@code Cell}
     */
    private Cell cell;

    /**
     * The index of the Cell in the Row
     */
    private Integer index;

    /**
     * Get to an ExcelCell instance from Apache POI Cell
     * @param cell The Cell instance to wrap
     * @return The ExcelCell instance
     * @since 0.5.0
     */
    public static ExcelCell of(Cell cell) {
        return ExcelCell.builder()
                .cell(cell)
                .index(cell.getColumnIndex())
                .build();
    }

    /**
     * Remove the selected Cell
     * @throws CellNotFoundException If the cell is not present or has not been created
     * @since 0.4.1
     */
    public void remove() throws CellNotFoundException {
        getRow().removeCell(this.index);
        this.cell = null;
        this.index = null;
    }

    /**
     * Returns the Row to which it belongs
     * @return A ExcelRow
     */
    public ExcelRow getRow() {
        Row row = this.cell.getRow();
        return ExcelRow.of(row);
    }

    /**
     * Read the value written inside the Cell
     * @return The value written in the Cell
     * @throws ReadValueException If an error occurs while reading
     * @since 0.4.0
     */
    public Object readValue() throws ReadValueException {
        Object val;
        switch (this.cell.getCellType()) {
            case BOOLEAN -> val = this.cell.getBooleanCellValue();
            case STRING -> val = this.cell.getStringCellValue();
            case NUMERIC -> val = DateUtil.isCellDateFormatted(this.cell)
                    ? this.cell.getDateCellValue()
                    : this.cell.getNumericCellValue();
            case FORMULA -> val = this.cell.getCellFormula();
            case BLANK -> val = "";
            default -> throw new ReadValueException("An error occurred while reading. CellType '" + this.cell.getCellType() + "'");
        }

        return val;
    }

    /**
     * Read the value written inside the Cell
     * @param type The class type of the object written to the Cell
     * @return The value written in the Cell
     * @throws ReadValueException If an error occurs while reading
     */
    public Object readValue(Class<?> type) throws ReadValueException {
        DataFormatter formatter = new DataFormatter(true);

        if(String.class.equals(type)) {
            return formatter.formatCellValue(this.cell);
        } else if (Boolean.class.equals(type)) {
            switch (this.cell.getCellType()) {
                case BOOLEAN -> {
                    return this.cell.getBooleanCellValue();
                }
                case FORMULA -> {
                    ExcelWorkbook excelWorkbook = this.getRow().getSheet().getWorkbook();
                    FormulaEvaluator formulaEvaluator = excelWorkbook.getFormulaEvaluator();
                    return formulaEvaluator.evaluate(this.cell).getBooleanValue();
                }
                default -> throw new ReadValueException("This type '" + type + "' is either incompatible with the CellType '" + this.cell.getCellType() + "'");
            }
        } else if (Integer.class.equals(type)) {
            if (this.cell.getCellType() == CellType.NUMERIC && !DateUtil.isCellDateFormatted(this.cell)) {
                return (int) this.cell.getNumericCellValue();
            }
            throw new ReadValueException("This type '" + type + "' is either incompatible with the CellType '" + this.cell.getCellType() + "'");
        } else if (Double.class.equals(type) && !DateUtil.isCellDateFormatted(this.cell)) {
            if (this.cell.getCellType() == CellType.NUMERIC) {
                return this.cell.getNumericCellValue();
            }
            throw new ReadValueException("This type '" + type + "' is either incompatible with the CellType '" + this.cell.getCellType() + "'");
        } else if (Long.class.equals(type)) {
            if (this.cell.getCellType() == CellType.NUMERIC && !DateUtil.isCellDateFormatted(this.cell)) {
                return (long) this.cell.getNumericCellValue();
            }
            throw new ReadValueException("This type '" + type + "' is either incompatible with the CellType '" + this.cell.getCellType() + "'");
        } else if (Date.class.equals(type)) {
            if (this.cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(this.cell)) {
                return this.cell.getDateCellValue();
            }
            throw new ReadValueException("This type '" + type + "' is either incompatible with the CellType '" + this.cell.getCellType() + "'");
        } else if (LocalDateTime.class.equals(type)) {
            if (this.cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(this.cell)) {
                return this.cell.getLocalDateTimeCellValue();
            }
            throw new ReadValueException("This type '" + type + "' is either incompatible with the CellType '" + this.cell.getCellType() + "'");
        } else if (LocalDate.class.equals(type)) {
            if (this.cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(this.cell)) {
                return this.cell.getLocalDateTimeCellValue().toLocalDate();
            }
            throw new ReadValueException("This type '" + type + "' is either incompatible with the CellType '" + this.cell.getCellType() + "'");
        } else {
            throw new ReadValueException("Cell type not supported");
        }
    }

    /**
     * Read the value written inside the Cell as String
     * @return The value written in the Cell
     * @since 0.4.0
     */
    public String readValueAsString() {
        DataFormatter formatter = new DataFormatter(true);
        return formatter.formatCellValue(this.cell);
    }

    /**
     * Writes inside the cell
     * @param val The value to write in the Cell
     */
    public void writeValue(Object val) {
        if (val == null) {
            this.cell.setCellValue("");
        } else if (val instanceof Integer || val instanceof Long) {
            this.formatStyle("0");
            this.cell.setCellValue(Integer.parseInt(String.valueOf(val)));
        } else if (val instanceof Double) {
            this.formatStyle("#0.00");
            this.cell.setCellValue(Double.parseDouble(String.valueOf(val)));
        } else if (val instanceof Date) {
            this.formatStyle("yyyy-MM-dd HH:mm");
            this.cell.setCellValue((Date) val);
        } else if (val instanceof LocalDate) {
            this.formatStyle("yyyy-MM-dd");
            this.cell.setCellValue((LocalDate) val);
        } else if (val instanceof LocalDateTime) {
            this.formatStyle("yyyy-MM-dd HH:mm");
            this.cell.setCellValue((LocalDateTime) val);
        } else if (val instanceof Boolean) {
            cell.setCellValue((Boolean) val);
        } else {
            cell.setCellValue(String.valueOf(val));
        }
    }

    /**
     * Returns the cell name
     * @return cell name
     * @since 0.4.2
     */
    public String getCellName() {
        int row = this.cell.getRowIndex();
        int col = this.cell.getColumnIndex();
        return ExcelUtility.getCellName(row, col);
    }

    /**
     * Format text according to the pattern provided
     * @param dataFormat The Apache POI library CellStyle dataFormat
     */
    public void formatStyle(short dataFormat) {
        ExcelWorkbook excelWorkbook = this.getRow().getSheet().getWorkbook();
        CellStyle newCellStyle = excelWorkbook.getWorkbook().createCellStyle();
        newCellStyle.cloneStyleFrom(this.cell.getCellStyle());
        newCellStyle.setDataFormat(dataFormat);
        this.cell.setCellStyle(newCellStyle);
    }

    /**
     * Format the cell according to the chosen format
     * @param dataFormat The string that defines the data format
     * @since 0.5.0
     */
    public void formatStyle(String dataFormat) {
        ExcelWorkbook excelWorkbook = this.getRow().getSheet().getWorkbook();
        CreationHelper creationHelper = excelWorkbook.getWorkbook().getCreationHelper();
        CellStyle newCellStyle = excelWorkbook.getWorkbook().createCellStyle();
        newCellStyle.cloneStyleFrom(this.cell.getCellStyle());
        newCellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(dataFormat));
        this.cell.setCellStyle(newCellStyle);
    }
}
