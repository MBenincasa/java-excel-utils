package samples.convertExcelFileToCsvFile;

import tools.implementations.ExcelConverterImpl;
import tools.interfaces.ExcelConverter;

import java.io.File;

public class Main {

    public static void main(String[] args) {

        ExcelConverter excelConverter = new ExcelConverterImpl();
        File excelFile = new File("./src/main/resources/employee.xlsx");

        try {
            System.out.println("Start the conversion...");
            File csvFile = excelConverter.excelToCsv(excelFile, "./src/main/resources/", "employee", "Employee");
            System.out.println("The file is ready. Path: " + csvFile.getAbsolutePath());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}