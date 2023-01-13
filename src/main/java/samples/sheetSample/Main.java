package samples.sheetSample;

import org.apache.poi.ss.usermodel.Workbook;
import tools.implementations.ExcelWorkbookUtilsImpl;
import tools.interfaces.ExcelSheetUtils;
import tools.implementations.ExcelSheetUtilsImpl;
import tools.interfaces.ExcelWorkbookUtils;

import java.io.File;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
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

            Workbook workbook = excelWorkbookUtils.open(file);
            String sheetNameTest = "test";
            int sheetIndexTest = 0;
            Boolean isPresentByName = excelSheetUtils.isPresent(workbook, sheetNameTest);
            System.out.println("Sheet is: " + sheetNameTest + " is present: " + isPresentByName);
            Boolean isPresentByPosition = excelSheetUtils.isPresent(workbook, 0);
            System.out.println("Sheet index: " + sheetIndexTest + " is present: " + isPresentByPosition);
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
