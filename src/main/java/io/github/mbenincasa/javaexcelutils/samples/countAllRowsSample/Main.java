package io.github.mbenincasa.javaexcelutils.samples.countAllRowsSample;

import io.github.mbenincasa.javaexcelutils.model.ExcelSheet;
import io.github.mbenincasa.javaexcelutils.model.ExcelWorkbook;

import java.io.File;
import java.util.LinkedList;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        File file = new File("./src/main/resources/car.xlsx");

        try {
            ExcelWorkbook excelWorkbook = ExcelWorkbook.open(file);
            List<ExcelSheet> excelSheets = excelWorkbook.getSheets();
            int totalRows = excelWorkbook.getSheet("car").countAllRows(true);
            System.out.println("Total rows: " + totalRows);
            List<Integer> totalRowsWithoutEmptyOfSheets = new LinkedList<>();
            for (ExcelSheet excelSheet : excelSheets) {
                totalRowsWithoutEmptyOfSheets.add(excelSheet.countAllRows(false));
            }

            System.out.println("Total of all sheets without empty rows: " + totalRowsWithoutEmptyOfSheets);
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }

    }
}
