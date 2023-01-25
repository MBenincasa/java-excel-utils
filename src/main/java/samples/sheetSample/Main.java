package samples.sheetSample;

import org.apache.poi.ss.usermodel.Workbook;
import tools.SheetUtility;
import tools.WorkbookUtility;

import java.io.File;
import java.util.List;

public class Main {

    public static void main(String[] args) {
        File file = new File("./src/main/resources/employee.xlsx");

        try {
            int totalSheets = SheetUtility.length(file);
            System.out.println("Total: " + totalSheets);
            List<String> sheetnames = SheetUtility.getNames(file);
            System.out.println("Sheet names: " + sheetnames);
            int sheetIndex = SheetUtility.getIndex(file, "Employee");
            System.out.println("Sheet index: " + sheetIndex);
            String sheetName = SheetUtility.getName(file, 0);
            System.out.println("Sheet name: " + sheetName);

            Workbook workbook = WorkbookUtility.open(file);
            String sheetNameTest = "test";
            int sheetIndexTest = 0;
            Boolean isPresentByName = SheetUtility.isPresent(workbook, sheetNameTest);
            System.out.println("Sheet is: " + sheetNameTest + ". It is present: " + isPresentByName);
            Boolean isPresentByPosition = SheetUtility.isPresent(workbook, 0);
            System.out.println("Sheet index: " + sheetIndexTest + ". It is present: " + isPresentByPosition);
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
