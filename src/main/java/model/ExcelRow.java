package model;

import lombok.AllArgsConstructor;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.LinkedList;
import java.util.List;

@AllArgsConstructor
@Getter
@EqualsAndHashCode
public class ExcelRow {

    private Row row;
    private Integer index;

    public List<ExcelCell> getCells() {
        List<ExcelCell> excelCells = new LinkedList<>();
        for (Cell cell : this.row) {
            excelCells.add(new ExcelCell(cell, cell.getColumnIndex()));
        }

        return excelCells;
    }

    @SneakyThrows
    public ExcelSheet getSheet() {
        Sheet sheet = this.row.getSheet();
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(sheet.getWorkbook());
        String sheetName = sheet.getSheetName();
        return new ExcelSheet(sheet, excelWorkbook.getSheet(sheetName).getIndex(), sheetName);
    }

    public Integer getLastColumnIndex() {
        return this.row.getLastCellNum() - 1;
    }

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
