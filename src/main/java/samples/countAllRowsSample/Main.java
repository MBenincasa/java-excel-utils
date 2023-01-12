package samples.countAllRowsSample;

import tools.interfaces.ExcelUtils;
import tools.implementations.ExcelUtilsImpl;

import java.io.File;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        ExcelUtils excelUtils = new ExcelUtilsImpl();
        File file = new File("./src/main/resources/car.xlsx");

        try {
            int totalRows = excelUtils.countAllRows(file, "car");
            System.out.println("Total: " + totalRows);
            int totalRowsWithoutEmpty = excelUtils.countAllRows(file, "car", false);
            System.out.println("Total without empty rows: " + totalRowsWithoutEmpty);
            List<Integer> totalRowsOfSheets = excelUtils.countAllRowsOfAllSheets(file);
            System.out.println("Total of all sheets: " + totalRowsOfSheets);
            List<Integer> totalRowsOfSheetsWithoutEmpty = excelUtils.countAllRowsOfAllSheets(file, false);
            System.out.println("Total of all sheets without empty rows : " + totalRowsOfSheetsWithoutEmpty);
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }

    }
}
