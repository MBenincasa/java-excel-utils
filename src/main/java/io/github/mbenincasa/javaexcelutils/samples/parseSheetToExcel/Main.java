package io.github.mbenincasa.javaexcelutils.samples.parseSheetToExcel;

import io.github.mbenincasa.javaexcelutils.model.excel.ExcelSheet;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelWorkbook;
import io.github.mbenincasa.javaexcelutils.model.parser.Direction;
import io.github.mbenincasa.javaexcelutils.model.parser.ExcelListParserMapping;

import java.io.File;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        File file = new File("./src/main/resources/parse_to_object.xlsx");
        try {
            ExcelWorkbook excelWorkbook = ExcelWorkbook.open(file);
            ExcelSheet excelSheet = excelWorkbook.getSheet("DATA");
            System.out.println("Start the parsing...");
            Employee employee = excelSheet.parseToObject(Employee.class, "A1");
            System.out.println("...completed");
            ExcelSheet excelSheet1 = excelWorkbook.getSheet("DATA_2");
            ExcelListParserMapping mapping = new ExcelListParserMapping("A1", Direction.VERTICAL, 8);
            List<Employee> employees = excelSheet1.parseToList(Employee.class, mapping);
            System.out.println("Data single object: " + employee.toString());
            System.out.println("Data multi objects: " + employees.toString());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }

    }
}
