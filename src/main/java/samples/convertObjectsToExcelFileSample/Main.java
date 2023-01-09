package samples.convertObjectsToExcelFileSample;

import enums.ExcelExtension;
import tools.interfaces.ExcelConverter;
import tools.implementations.ExcelConverterImpl;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        ExcelConverter excelConverter = new ExcelConverterImpl();
        List<Employee> employees = new ArrayList<>();
        employees.add(new Employee("Mario", "Rossi", 25, new Date(), 28000.00));
        employees.add(new Employee("Giuseppe", "Verdi", 22, new Date(), 23670.89));

        try {
            System.out.println("Start the conversation...");
            File report = excelConverter.convertObjectsToExcelFile(employees, Employee.class, "./src/main/resources/", "employee", ExcelExtension.XLSX, true);
            System.out.println("The file is ready. Path: " + report.getAbsolutePath());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
