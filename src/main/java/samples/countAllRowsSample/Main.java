package samples.countAllRowsSample;

import tools.ExcelUtils;
import tools.ExcelUtilsImpl;

import java.io.File;

public class Main {

    public static void main(String[] args) {

        ExcelUtils excelUtils = new ExcelUtilsImpl();
        File file = new File("./src/main/resources/employee.xlsx");

        try {
            int totalRows = excelUtils.countAllRows(file, true);
            System.out.println("Total: " + totalRows);
            int totalRowsWithoutEmpty = excelUtils.countAllRows(file, false, "Employee");
            System.out.println("Total without empty rows: " + totalRowsWithoutEmpty);
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }

    }
}
