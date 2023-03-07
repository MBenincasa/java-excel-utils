package io.github.mbenincasa.javaexcelutils.samples.convertExcelFileToObjectsSample;

import io.github.mbenincasa.javaexcelutils.annotations.ExcelBodyStyle;
import io.github.mbenincasa.javaexcelutils.annotations.ExcelField;
import io.github.mbenincasa.javaexcelutils.annotations.ExcelHeaderStyle;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.ToString;

@AllArgsConstructor
@NoArgsConstructor
@Setter
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
