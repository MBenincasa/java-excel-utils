package samples.sheetSample;

import tools.interfaces.ExcelSheetUtils;
import tools.implementations.ExcelSheetUtilsImpl;

import java.io.File;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        ExcelSheetUtils excelSheetUtils = new ExcelSheetUtilsImpl();
        File file = new File("./src/main/resources/employee.xlsx");

        try {
            int totalSheets = excelSheetUtils.countAll(file);
            System.out.println("Total: " + totalSheets);
            List<String> sheetnames = excelSheetUtils.getAllNames(file);
            System.out.println("Sheet names: " + sheetnames.toString());
            int sheetIndex = excelSheetUtils.getIndex(file, "Employee");
            System.out.println("Sheet index: " + sheetIndex);
            String sheetName = excelSheetUtils.getNameByIndex(file, 0);
            System.out.println("Sheet name: " + sheetName);
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
