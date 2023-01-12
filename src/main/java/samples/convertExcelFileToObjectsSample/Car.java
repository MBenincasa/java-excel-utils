package samples.convertExcelFileToObjectsSample;

import annotations.ExcelBodyStyle;
import annotations.ExcelField;
import annotations.ExcelHeaderStyle;
import lombok.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

@AllArgsConstructor
@NoArgsConstructor
@Setter
@ToString
@ExcelHeaderStyle(cellColor = IndexedColors.ORANGE, horizontal = HorizontalAlignment.RIGHT, vertical = VerticalAlignment.BOTTOM, autoSize = true)
@ExcelBodyStyle(cellColor = IndexedColors.LIGHT_ORANGE, horizontal = HorizontalAlignment.RIGHT, vertical = VerticalAlignment.BOTTOM)
public class Car {

    @ExcelField(name = "BRAND")
    private String brand;
    @ExcelField(name = "MODEL")
    private String model;
    @ExcelField(name = "YEAR")
    private Integer year;
    @ExcelField(name = "RELEASE DATE")
    private LocalDate releaseDate;
    @ExcelField(name = "FIRST SALE")
    private Date firstSale;
    @ExcelField(name = "LAST SALE")
    private LocalDateTime lastSale;
}
