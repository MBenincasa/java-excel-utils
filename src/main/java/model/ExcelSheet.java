package model;

import exceptions.SheetAlreadyExistsException;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

@AllArgsConstructor
@Getter
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

    public ExcelWorkbook getExcelWorkbook() {
        return new ExcelWorkbook(this.getSheet().getWorkbook());
    }
}
