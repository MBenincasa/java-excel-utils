package samples.convertObjectsToExcelFile;

import utils.converter.ExcelConverter;
import utils.converter.ExcelConverterImpl;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        ExcelConverter excelConverter = new ExcelConverterImpl();
        List<Employee> employees = new ArrayList<>();
        employees.add(new Employee("Mario", "Rossi", 27, 28000.00));
        employees.add(new Employee("Giuseppe", "Verdi", 22, 23670.89));

        try {
            File report = excelConverter.convertObjectsToExcelFile(employees, Employee.class, "./src/main/resources/test.xlsx", true);
        } catch (IllegalAccessException | IOException e) {
            throw new RuntimeException(e);
        }
    }
}
