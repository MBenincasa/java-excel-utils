package io.github.mbenincasa.javaexcelutils.tools.utils;

import io.github.mbenincasa.javaexcelutils.annotations.ExcelCellMapping;
import io.github.mbenincasa.javaexcelutils.model.parser.ExcelCellParser;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.ToString;

import java.time.LocalDate;

@AllArgsConstructor
@NoArgsConstructor
@Getter
@ToString
public class ParsableEmployee implements ExcelCellParser {

    @ExcelCellMapping(deltaRow = 0, deltaCol = 1)
    private String name;
    @ExcelCellMapping(deltaRow = 1, deltaCol = 1)
    private String lastName;
    @ExcelCellMapping(deltaRow = 2, deltaCol = 1)
    private Integer age;
    @ExcelCellMapping(deltaRow = 3, deltaCol = 1)
    private LocalDate hireDate;
    @ExcelCellMapping(deltaRow = 3, deltaCol = 2)
    private LocalDate terminationDate;
    @ExcelCellMapping(deltaRow = 4, deltaCol = 4)
    private ParsableAddress address;
}
