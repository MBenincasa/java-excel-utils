package io.github.mbenincasa.javaexcelutils.samples.parseSheetToExcel;

import io.github.mbenincasa.javaexcelutils.annotations.ExcelCellMapping;
import io.github.mbenincasa.javaexcelutils.model.parser.ExcelCellParser;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.ToString;

@AllArgsConstructor
@NoArgsConstructor
@Setter
@ToString
public class Address implements ExcelCellParser {

    @ExcelCellMapping(deltaRow = 1, deltaCol = 0)
    private String city;
    @ExcelCellMapping(deltaRow = 2, deltaCol = 0)
    private String cap;
}
