package io.github.mbenincasa.javaexcelutils.samples.convertExcelFileToObjects;

import io.github.mbenincasa.javaexcelutils.model.converter.ExcelToObject;
import io.github.mbenincasa.javaexcelutils.tools.Converter;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

public class Main {

    public static void main(String[] args) {

        File file = new File("./src/main/resources/result.xlsx");
        ExcelToObject<Employee> employeeExcelToObject = new ExcelToObject<>("Employee", Employee.class);
        ExcelToObject<Office> officeExcelToObject = new ExcelToObject<>("Office", Office.class);
        List<ExcelToObject<?>> excelToObjects = new ArrayList<>();
        excelToObjects.add(employeeExcelToObject);
        excelToObjects.add(officeExcelToObject);

        try {
            System.out.println("Start the conversion...");
            Map<String, Stream<?>> map = Converter.excelFileToObjects(file, excelToObjects);
            System.out.println("...completed");
            for (Map.Entry<String, Stream<?>> entry : map.entrySet()) {
                System.out.println("Sheet: " + entry.getKey());
                System.out.println("Data: " + entry.getValue().toList());
            }
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
