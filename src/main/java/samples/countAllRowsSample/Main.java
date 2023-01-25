package samples.countAllRowsSample;

import tools.ExcelUtility;

import java.io.File;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        File file = new File("./src/main/resources/car.xlsx");

        try {
            int totalRows = ExcelUtility.countAllRows(file, "car");
            System.out.println("Total: " + totalRows);
            int totalRowsWithoutEmpty = ExcelUtility.countAllRows(file, "car", false);
            System.out.println("Total without empty rows: " + totalRowsWithoutEmpty);
            List<Integer> totalRowsOfSheets = ExcelUtility.countAllRowsOfAllSheets(file);
            System.out.println("Total of all sheets: " + totalRowsOfSheets);
            List<Integer> totalRowsOfSheetsWithoutEmpty = ExcelUtility.countAllRowsOfAllSheets(file, false);
            System.out.println("Total of all sheets without empty rows : " + totalRowsOfSheetsWithoutEmpty);
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }

    }
}
