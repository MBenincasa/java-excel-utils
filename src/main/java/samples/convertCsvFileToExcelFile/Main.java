package samples.convertCsvFileToExcelFile;

import enums.Extension;
import org.apache.poi.ss.usermodel.Workbook;
import tools.Converter;
import tools.WorkbookUtility;

import java.io.File;
import java.io.FileOutputStream;

public class Main {

    public static void main(String[] args) {

        File csvFile = new File("./src/main/resources/employee.csv");
        File csvFile2 = new File("./src/main/resources/employee_2.csv");

        try {
            System.out.println("Start the conversion...");
            File excelFile = Converter.csvToExcel(csvFile, "./src/main/resources/", "employee_2", Extension.XLSX);
            System.out.println("First conversion completed...");

            Workbook workbook = WorkbookUtility.open(excelFile);
            Converter.csvToExistingExcel(workbook, csvFile2);
            FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
            workbook.write(fileOutputStream);
            WorkbookUtility.close(workbook, fileOutputStream);
            System.out.println("The file is ready. Path: " + excelFile.getAbsolutePath());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
