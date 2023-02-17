package tools.utils;

import annotations.ExcelBodyStyle;
import annotations.ExcelField;
import annotations.ExcelHeaderStyle;
import lombok.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

@AllArgsConstructor
@NoArgsConstructor
@ToString
@Getter
@Setter
@ExcelHeaderStyle(cellColor = IndexedColors.ORANGE, horizontal = HorizontalAlignment.RIGHT, vertical = VerticalAlignment.BOTTOM, autoSize = true)
@ExcelBodyStyle(cellColor = IndexedColors.LIGHT_ORANGE, horizontal = HorizontalAlignment.RIGHT, vertical = VerticalAlignment.BOTTOM)
public class Person {

    @ExcelField(name = "LAST NAME")
    private String lastName;
    @ExcelField(name = "NAME")
    private String name;
    @ExcelField(name = "AGE")
    private Integer age;
}
