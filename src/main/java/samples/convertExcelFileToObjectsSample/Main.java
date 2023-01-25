package samples.convertExcelFileToObjectsSample;

import tools.Converter;

import java.io.File;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        File file = new File("./src/main/resources/car.xlsx");

        try {
            System.out.println("Start the conversion...");
            List<Car> employees = (List<Car>) Converter.excelToObjects(file, Car.class);
            System.out.println("The list is ready. List: " + employees.toString());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
