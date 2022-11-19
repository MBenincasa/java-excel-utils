package samples.countAllRows;

import tools.utils.ExcelUtils;
import tools.utils.ExcelUtilsImpl;

import java.io.File;

public class Main {

    public static void main(String[] args) {

        ExcelUtils excelUtils = new ExcelUtilsImpl();
        File file = new File("./src/main/resources/employee.xlsx");
        try {
            int totalRows = excelUtils.countAllRows(file, true);
            System.out.println("Total: " + totalRows);
            int totalRowsWithoutEmpty = excelUtils.countAllRows(file, false);
            System.out.println("Total without empty rows: " + totalRowsWithoutEmpty);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }
}
