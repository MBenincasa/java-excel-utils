package samples.sheetSample;

import tools.ExcelUtils;
import tools.ExcelUtilsImpl;

import java.io.File;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        ExcelUtils excelUtils = new ExcelUtilsImpl();
        File file = new File("./src/main/resources/employee.xlsx");

        try {
            int totalSheets = excelUtils.countAllSheets(file);
            System.out.println("Total: " + totalSheets);
            List<String> sheetnames = excelUtils.getAllSheetNames(file);
            System.out.println("Sheet names: " + sheetnames.toString());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
