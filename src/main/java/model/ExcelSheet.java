package model;

import exceptions.SheetAlreadyExistsException;
import lombok.AllArgsConstructor;
import lombok.EqualsAndHashCode;
import lombok.Getter;
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
}
