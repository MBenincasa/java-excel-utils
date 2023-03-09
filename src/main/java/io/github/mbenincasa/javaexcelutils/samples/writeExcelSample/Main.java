package io.github.mbenincasa.javaexcelutils.samples.writeExcelSample;

import io.github.mbenincasa.javaexcelutils.model.excel.ExcelCell;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelRow;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelSheet;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelWorkbook;
import org.apache.commons.io.FilenameUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

public class Main {

    public static void main(String[] args) {
        File testFile = new File("./src/main/resources/test.xlsx");

        try {
            ExcelWorkbook excelWorkbook = ExcelWorkbook.create(FilenameUtils.getExtension(testFile.getName()));
            ExcelSheet excelSheet = excelWorkbook.createSheet("TEST");
            ExcelRow excelRow = excelSheet.createRow(0);
            ExcelCell excelCell = excelRow.createCell(0);
            excelCell.writeValue("Rossi");
            ExcelCell excelCell1 = excelRow.createCell(1);
            excelCell1.writeValue("Mario");
            ExcelCell excelCell2 = excelRow.createCell(2);
            excelCell2.writeValue(23);
            ExcelCell excelCell3 = excelRow.createCell(3);
            excelCell3.writeValue(LocalDateTime.now());
            ExcelCell excelCell4 = excelRow.createCell(4);
            excelCell4.writeValue(LocalDate.now());
            ExcelCell excelCell5 = excelRow.createCell(5);
            excelCell5.writeValue(new Date());
            ExcelCell excelCell6 = excelRow.createCell(6);
            excelCell6.writeValue(21.12);
            ExcelCell excelCell7 = excelRow.createCell(7);
            excelCell7.writeValue(true);

            FileOutputStream fileOutputStream = new FileOutputStream(testFile);
            excelWorkbook.getWorkbook().write(fileOutputStream);
            excelWorkbook.close(fileOutputStream);
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage());
        }

    }
}
