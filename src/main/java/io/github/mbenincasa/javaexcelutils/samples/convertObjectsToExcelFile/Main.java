package io.github.mbenincasa.javaexcelutils.samples.convertObjectsToExcelFile;

import io.github.mbenincasa.javaexcelutils.enums.Extension;
import io.github.mbenincasa.javaexcelutils.model.converter.ObjectToExcel;
import io.github.mbenincasa.javaexcelutils.tools.Converter;

import java.io.File;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Stream;

public class Main {

    public static void main(String[] args) {

        List<Employee> employees = new ArrayList<>();
        employees.add(new Employee("Rossi", "Mario", 25, LocalDate.of(1987, 5, 22), new Date(), 28000.00, LocalDateTime.now(), true));
        employees.add(new Employee("Verdi", "Giuseppe", 22, LocalDate.of(1991, 2, 23), new Date(), 23670.89, LocalDateTime.now(), false));

        List<Office> offices = new LinkedList<>();
        offices.add(new Office("Nocera Inferiore", "Salerno", 40));
        offices.add(new Office("Pero", "Milano", 73));

        try {
            Stream<Employee> employeeStream = employees.stream();
            Stream<Office> officeStream = offices.stream();
            List<ObjectToExcel<?>> list = new ArrayList<>();
            list.add(new ObjectToExcel<>("Employee", Employee.class, employeeStream));
            list.add(new ObjectToExcel<>("Office", Office.class, officeStream));
            System.out.println("Converting...");
            File fileOutput = Converter.objectsToExcelFile(list, Extension.XLSX, "./src/main/resources/result", true);
            System.out.println("...completed");
            System.out.println("File output: " + fileOutput.getAbsolutePath());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
