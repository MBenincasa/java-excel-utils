package io.github.mbenincasa.javaexcelutils.samples.convertExcelFileToJsonFile;

import io.github.mbenincasa.javaexcelutils.tools.Converter;

import java.io.File;

public class Main {

    public static void main(String[] args) {

        File excelFile = new File("./src/main/resources/employee.xlsx");

        try {
            System.out.println("Start the conversion...");
            File jsonFile = Converter.excelToJsonFile(excelFile, "./src/main/resources/result");
            System.out.println("... completed");
            System.out.println("The file is ready. Path: " + jsonFile.getAbsolutePath());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
