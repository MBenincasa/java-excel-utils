package samples.convertExcelFileToObjectsSample;

import tools.implementations.ExcelConverterImpl;
import tools.interfaces.ExcelConverter;

import java.io.File;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        File file = new File("./src/main/resources/car.xlsx");

        try {
            System.out.println("Start the conversion...");
            ExcelConverter excelConverter = new ExcelConverterImpl();
            List<Car> employees = (List<Car>) excelConverter.excelToObjects(file, Car.class);
            System.out.println("The list is ready. List: " + employees.toString());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
