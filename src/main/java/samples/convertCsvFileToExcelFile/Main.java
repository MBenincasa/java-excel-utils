package samples.convertCsvFileToExcelFile;

import enums.Extension;
import tools.implementations.ExcelConverterImpl;
import tools.interfaces.ExcelConverter;

import java.io.File;

public class Main {

    public static void main(String[] args) {

        ExcelConverter excelConverter = new ExcelConverterImpl();
        File csvFile = new File("./src/main/resources/employee.csv");

        try {
            System.out.println("Start the conversion...");
            File excelFile = excelConverter.csvToExcel(csvFile, "./src/main/resources/", "employee_2", Extension.XLSX);
            System.out.println("The file is ready. Path: " + excelFile.getAbsolutePath());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
