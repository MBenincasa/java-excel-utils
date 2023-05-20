package io.github.mbenincasa.javaexcelutils.samples.convertObjectsToExcelFile;

import io.github.mbenincasa.javaexcelutils.annotations.ExcelBodyStyle;
import io.github.mbenincasa.javaexcelutils.annotations.ExcelField;
import io.github.mbenincasa.javaexcelutils.annotations.ExcelHeaderStyle;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import lombok.ToString;

@AllArgsConstructor
@NoArgsConstructor
@ToString
@ExcelHeaderStyle(autoSize = true)
@ExcelBodyStyle
public class Office {

    @ExcelField(name = "CITY")
    private String city;
    @ExcelField(name = "PROVINCE")
    private String province;
    @ExcelField(name = "NUMBER OF STATIONS")
    private Integer numStations;
}
