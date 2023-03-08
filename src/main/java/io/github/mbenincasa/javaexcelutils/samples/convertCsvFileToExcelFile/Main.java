package io.github.mbenincasa.javaexcelutils.samples.convertCsvFileToExcelFile;

import io.github.mbenincasa.javaexcelutils.enums.Extension;
import io.github.mbenincasa.javaexcelutils.tools.Converter;

import java.io.File;

public class Main {

    public static void main(String[] args) {

        File csvFile = new File("./src/main/resources/Employee.csv");

        try {
            System.out.println("Start the conversion...");
            File excelFile = Converter.csvToExcelFile(csvFile, "Employee", "./src/main/resources/employee_2", Extension.XLSX);
            System.out.println("... completed");
            System.out.println("The file is ready. Path: " + excelFile.getAbsolutePath());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
