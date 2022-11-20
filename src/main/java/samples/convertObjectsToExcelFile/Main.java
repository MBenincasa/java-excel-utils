package samples.convertObjectsToExcelFile;

import enums.Extension;
import tools.ExcelConverter;
import tools.ExcelConverterImpl;

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
            System.out.println("Start the conversation...");
            File report = excelConverter.convertObjectsToExcelFile(employees, Employee.class, "./src/main/resources/", "employee", Extension.XLSX, true);
            System.out.println("The file is ready");
        } catch (IllegalAccessException | IOException e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
