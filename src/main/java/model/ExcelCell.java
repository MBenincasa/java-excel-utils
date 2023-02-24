package model;

import lombok.AllArgsConstructor;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

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
}
