package samples.convertObjectsToExcelFile;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.ToString;

@Data
@ToString
@AllArgsConstructor
public class Employee {

    private String name;
    private String lastName;
    private Integer age;
    private Double salary;
}
