package io.github.mbenincasa.javaexcelutils.tools.utils;

import io.github.mbenincasa.javaexcelutils.annotations.ExcelCellMapping;
import io.github.mbenincasa.javaexcelutils.model.parser.ExcelCellParser;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.ToString;

@AllArgsConstructor
@NoArgsConstructor
@Getter
@ToString
public class ParsableAddress implements ExcelCellParser {

    @ExcelCellMapping(deltaRow = 1, deltaCol = 0)
    private String city;
    @ExcelCellMapping(deltaRow = 2, deltaCol = 0)
    private String cap;
}
