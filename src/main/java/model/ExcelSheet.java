package model;

import exceptions.SheetAlreadyExistsException;
import lombok.AllArgsConstructor;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.LinkedList;
import java.util.List;

@AllArgsConstructor
@Getter
@EqualsAndHashCode
public class ExcelSheet {

    private Sheet sheet;
    private Integer index;
    private String name;

    public static ExcelSheet create(ExcelWorkbook excelWorkbook) throws SheetAlreadyExistsException {
        return create(excelWorkbook, null);
    }

    public static ExcelSheet create(ExcelWorkbook excelWorkbook, String sheetName) throws SheetAlreadyExistsException {
        Workbook workbook = excelWorkbook.getWorkbook();
        Sheet sheet;
        try {
            sheet = (sheetName == null || sheetName.isEmpty())
                    ? workbook.createSheet()
                    : workbook.createSheet(sheetName);
        } catch (IllegalArgumentException e) {
            throw new SheetAlreadyExistsException(e.getMessage(), e.getCause());
        }
        return new ExcelSheet(sheet, workbook.getSheetIndex(sheet), sheet.getSheetName());
    }

    public ExcelWorkbook getWorkbook() {
        return new ExcelWorkbook(this.getSheet().getWorkbook());
    }

    public List<ExcelRow> getRows() {
        List<ExcelRow> excelRows = new LinkedList<>();
        for (Row row : this.sheet) {
            excelRows.add(new ExcelRow(row, row.getRowNum()));
        }
        
        return excelRows;
    }

    public Integer getLastRowIndex() {
        return this.sheet.getLastRowNum();
    }

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
