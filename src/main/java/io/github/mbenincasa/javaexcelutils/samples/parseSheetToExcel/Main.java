package io.github.mbenincasa.javaexcelutils.samples.parseSheetToExcel;

import io.github.mbenincasa.javaexcelutils.model.excel.ExcelSheet;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelWorkbook;

import java.io.File;

public class Main {

    public static void main(String[] args) {

        File file = new File("./src/main/resources/parse_to_object.xlsx");
        try {
            ExcelWorkbook excelWorkbook = ExcelWorkbook.open(file);
            ExcelSheet excelSheet = excelWorkbook.getSheet("DATA");
            System.out.println("Start the parsing...");
            Employee employee = excelSheet.parseToObject(Employee.class, "A1");
            System.out.println("...completed");
            System.out.println("Data: " + employee.toString());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }

    }
}
