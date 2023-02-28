package model;

import exceptions.ReadValueException;
import lombok.AllArgsConstructor;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

@AllArgsConstructor
@Getter
@EqualsAndHashCode
public class ExcelCell {

    private Cell cell;
    private Integer index;

    public ExcelRow getRow() {
        Row row = this.cell.getRow();
        return new ExcelRow(row, row.getRowNum());
    }

    public Object readValue(Class<?> type) throws ReadValueException {
        Object val;
        switch (this.cell.getCellType()) {
            case BOOLEAN -> val = this.cell.getBooleanCellValue();
            case STRING -> val = this.cell.getStringCellValue();
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
                    throw new ReadValueException("This numeric type is not supported: " + type);
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

    public void formatStyle(short dataFormat) {
        ExcelWorkbook excelWorkbook = this.getRow().getSheet().getWorkbook();
        CellStyle newCellStyle = excelWorkbook.getWorkbook().createCellStyle();
        newCellStyle.cloneStyleFrom(this.cell.getCellStyle());
        newCellStyle.setDataFormat(dataFormat);
        this.cell.setCellStyle(newCellStyle);
    }
}
