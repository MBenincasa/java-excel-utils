package samples.convertObjectsToExcelFileSample;

import enums.Extension;
import tools.Converter;

import java.io.File;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        List<Employee> employees = new ArrayList<>();
        employees.add(new Employee("Rossi", "Mario", 25, LocalDate.of(1987, 5, 22), new Date(), 28000.00, LocalDateTime.now(), true));
        employees.add(new Employee("Verdi", "Giuseppe", 22, LocalDate.of(1991, 2, 23), new Date(), 23670.89, LocalDateTime.now(), false));

        try {
            System.out.println("Start the conversion...");
            File report = Converter.objectsToExcel(employees, Employee.class, "./src/main/resources/", "employee", Extension.XLSX, true);
            System.out.println("The file is ready. Path: " + report.getAbsolutePath());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
