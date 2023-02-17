package tools.utils;

import annotations.ExcelBodyStyle;
import annotations.ExcelField;
import annotations.ExcelHeaderStyle;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import lombok.ToString;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

@AllArgsConstructor
@NoArgsConstructor
@ToString
@ExcelHeaderStyle(cellColor = IndexedColors.ORANGE, horizontal = HorizontalAlignment.RIGHT, vertical = VerticalAlignment.BOTTOM, autoSize = true)
@ExcelBodyStyle(cellColor = IndexedColors.LIGHT_ORANGE, horizontal = HorizontalAlignment.RIGHT, vertical = VerticalAlignment.BOTTOM)
public class Address {

    @ExcelField(name = "CITY")
    private String city;
    @ExcelField(name = "ADDRESS")
    private String address;
}
