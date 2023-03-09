package io.github.mbenincasa.javaexcelutils.samples.convertJsonFileToExcelFile;

import io.github.mbenincasa.javaexcelutils.enums.Extension;
import io.github.mbenincasa.javaexcelutils.model.converter.JsonToExcel;
import io.github.mbenincasa.javaexcelutils.tools.Converter;

import java.io.File;

public class Main {

    public static void main(String[] args) {

        File jsonFile = new File("./src/main/resources/office.json");

        try {
            System.out.println("Start the conversion...");
            JsonToExcel<Office> officeJsonToExcel = new JsonToExcel<>("office", Office.class);
            File excelFile = Converter.jsonToExcelFile(jsonFile, officeJsonToExcel, Extension.XLSX, "./src/main/resources/from-json-to-excel", true);
            System.out.println("... completed");
            System.out.println("The file is ready. Path: " + excelFile.getAbsolutePath());
        } catch (Exception e) {
            System.err.println("There was an error. Check the console");
            throw new RuntimeException(e);
        }
    }
}
