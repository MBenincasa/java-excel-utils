package io.github.mbenincasa.javaexcelutils.samples.convertObjectsToExcelFileSample;

import io.github.mbenincasa.javaexcelutils.enums.Extension;
import io.github.mbenincasa.javaexcelutils.tools.Converter;

import java.io.File;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        List<Employee> employees = new ArrayList<>();
        employees.add(new Employee("Rossi", "Mario", 25, LocalDate.of(1987, 5, 22), new Date(), 28000.00, LocalDateTime.now(), true));
        employees.add(new Employee("Verdi", "Giuseppe", 22, LocalDate.of(1991, 2, 23), new Date(), 23670.89, LocalDateTime.now(), false));

        List<Office> offices = new LinkedList<>();
        offices.add(new Office("Nocera Inferiore", "Salerno", 40));
        offices.add(new Office("Pero", "Milano", 73));

        try {
            System.out.println("Start the conversion...");
            File report = Converter.objectsToExcel(employees, Employee.class, "./src/main/resources/", "employee", Extension.XLSX, true);
            System.out.println("First conversion completed...");
            Converter.objectsToExistingExcel(report, offices, Office.class, true);
            System.out.println("The file is ready. Path: " + report.getAbsolutePath());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
