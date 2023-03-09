package io.github.mbenincasa.javaexcelutils.model.excel;

import io.github.mbenincasa.javaexcelutils.exceptions.ReadValueException;
import lombok.AllArgsConstructor;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import org.apache.poi.ss.usermodel.*;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

/**
 * {@code ExcelCell} is the {@code Cell} wrapper class of the Apache POI library
 * @author Mirko Benincasa
 * @since 0.3.0
 */
@AllArgsConstructor
@Getter
@EqualsAndHashCode
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
     * Returns the Row to which it belongs
     * @return A ExcelRow
     */
    public ExcelRow getRow() {
        Row row = this.cell.getRow();
        return new ExcelRow(row, row.getRowNum());
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
            case NUMERIC -> val = this.cell.getNumericCellValue();
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
        Object val;
        switch (this.cell.getCellType()) {
            case BOOLEAN -> {
                if (!Boolean.class.equals(type)) {
                    throw new ReadValueException("This type '" + type + "' is either incompatible with the CellType '" + this.cell.getCellType() + "'");
                }
                val = this.cell.getBooleanCellValue();
            }
            case STRING -> {
                if (!String.class.equals(type)) {
                    throw new ReadValueException("This type '" + type + "' is either incompatible with the CellType '" + this.cell.getCellType() + "'");
                }
                val = this.cell.getStringCellValue();
            }
            case NUMERIC -> {
                if (Integer.class.equals(type)) {
                    val = (int) this.cell.getNumericCellValue();
                } else if (Double.class.equals(type)) {
                    val = this.cell.getNumericCellValue();
                } else if (Long.class.equals(type)) {
                    val = (long) this.cell.getNumericCellValue();
                } else if (Date.class.equals(type)) {
                    val = this.cell.getDateCellValue();
                } else if (LocalDateTime.class.equals(type)) {
                    val = this.cell.getLocalDateTimeCellValue();
                } else if (LocalDate.class.equals(type)) {
                    val = this.cell.getLocalDateTimeCellValue().toLocalDate();
                } else {
                    throw new ReadValueException("This type '" + type + "' is either incompatible with the CellType '" + this.cell.getCellType() + "' or not yet supported");
                }
            }
            case FORMULA -> {
                ExcelWorkbook excelWorkbook = this.getRow().getSheet().getWorkbook();
                FormulaEvaluator formulaEvaluator = excelWorkbook.getFormulaEvaluator();
                if (Boolean.class.equals(type)) {
                    val = formulaEvaluator.evaluate(this.cell).getBooleanValue();
                } else {
                    val = this.cell.getCellFormula();
                }
            }
            default -> throw new ReadValueException("Cell type not supported");
        }

        return val;
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
        if (val instanceof Integer || val instanceof Long) {
            this.formatStyle((short) 1);
            this.cell.setCellValue(Integer.parseInt(String.valueOf(val)));
        } else if (val instanceof Double) {
            this.formatStyle((short) 4);
            this.cell.setCellValue(Double.parseDouble(String.valueOf(val)));
        } else if (val instanceof Date) {
            this.formatStyle((short) 22);
            this.cell.setCellValue((Date) val);
        } else if (val instanceof LocalDate) {
            this.formatStyle((short) 14);
            this.cell.setCellValue((LocalDate) val);
        } else if (val instanceof LocalDateTime) {
            this.formatStyle((short) 22);
            this.cell.setCellValue((LocalDateTime) val);
        } else if (val instanceof Boolean) {
            cell.setCellValue((Boolean) val);
        } else {
            cell.setCellValue(String.valueOf(val));
        }
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
}
