package io.github.mbenincasa.javaexcelutils.samples.convertExcelFileToCsvFile;

import io.github.mbenincasa.javaexcelutils.tools.Converter;

import java.io.File;
import java.util.Map;

public class Main {

    public static void main(String[] args) {

        File excelFile = new File("./src/main/resources/employee.xlsx");

        try {
            System.out.println("Start the conversion...");
            Map<String, File> fileMap = Converter.excelToCsvFile(excelFile, "./src/main/resources");
            System.out.println("... completed");
            for (Map.Entry<String, File> entry : fileMap.entrySet()) {
                System.out.println("The file is ready. Path: " + entry.getValue().getAbsolutePath());
            }
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
