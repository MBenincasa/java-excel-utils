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
@Getter
@Setter
@ToString
@ExcelHeaderStyle(cellColor = IndexedColors.ORANGE, horizontal = HorizontalAlignment.RIGHT, vertical = VerticalAlignment.BOTTOM, autoSize = true)
@ExcelBodyStyle(cellColor = IndexedColors.LIGHT_ORANGE, horizontal = HorizontalAlignment.RIGHT, vertical = VerticalAlignment.BOTTOM)
public class Employee {

    @ExcelField(name = "LAST NAME")
    private String lastName;
    @ExcelField(name = "NAME")
    private String name;
    @ExcelField(name = "AGE")
    private Integer age;
    @ExcelField(name = "BIRTHDAY")
    private LocalDate birthday;
    @ExcelField(name = "HIRE DATE")
    private Date hireDate;
    @ExcelField(name = "SALARY (€)")
    private Double salary;
    @ExcelField(name = "LAST SIGN IN")
    private LocalDateTime lastSignIn;
    @ExcelField(name = "IS IN OFFICE")
    private Boolean isInOffice;
}
