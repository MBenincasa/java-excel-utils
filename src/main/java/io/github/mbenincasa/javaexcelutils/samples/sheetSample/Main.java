package io.github.mbenincasa.javaexcelutils.samples.sheetSample;

import io.github.mbenincasa.javaexcelutils.model.ExcelSheet;
import io.github.mbenincasa.javaexcelutils.model.ExcelWorkbook;

import java.io.File;
import java.util.List;

public class Main {

    public static void main(String[] args) {
        File file = new File("./src/main/resources/employee.xlsx");

        try {
            ExcelWorkbook excelWorkbook = ExcelWorkbook.open(file);
            int totalSheets = excelWorkbook.countSheets();
            System.out.println("Total: " + totalSheets);
            List<String> sheetnames = excelWorkbook.getSheets().stream().map(ExcelSheet::getName).toList();
            System.out.println("Sheet names: " + sheetnames);
            int sheetIndex = excelWorkbook.getSheet("Employee").getIndex();
            System.out.println("Sheet index: " + sheetIndex);
            String sheetName = excelWorkbook.getSheet(0).getName();
            System.out.println("Sheet name: " + sheetName);

            String sheetNameTest = "test";
            int sheetIndexTest = 0;
            Boolean isPresentByName = excelWorkbook.isSheetPresent(sheetNameTest);
            System.out.println("Sheet is: " + sheetNameTest + ". It is present: " + isPresentByName);
            Boolean isPresentByPosition = excelWorkbook.isSheetPresent(0);
            System.out.println("Sheet index: " + sheetIndexTest + ". It is present: " + isPresentByPosition);
            excelWorkbook.close();
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
