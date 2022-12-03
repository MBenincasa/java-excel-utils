package samples.convertObjectsToExcelFileSample;

import annotations.ExcelHeader;
import lombok.AllArgsConstructor;
import lombok.ToString;

@AllArgsConstructor
@ToString
public class Employee {

    @ExcelHeader(name = "NAME")
    private String name;
    @ExcelHeader(name = "LAST NAME")
    private String lastName;
    @ExcelHeader(name = "AGE")
    private Integer age;
    @ExcelHeader(name = "SALARY (€)")
    private Double salary;
}
