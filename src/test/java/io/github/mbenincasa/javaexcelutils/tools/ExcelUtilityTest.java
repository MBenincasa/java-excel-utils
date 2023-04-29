package io.github.mbenincasa.javaexcelutils.tools;

import io.github.mbenincasa.javaexcelutils.exceptions.ExtensionNotValidException;
import org.apache.commons.io.FilenameUtils;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;

public class ExcelUtilityTest {

    private final File excelFile = new File("./src/test/resources/employee_2.xlsx");

    @Test
    void checkExcelExtension() throws ExtensionNotValidException {
        String filename = excelFile.getName();
        String extension = ExcelUtility.checkExcelExtension(filename);
        Assertions.assertEquals("xlsx", extension);
    }

    @Test
    void isValidExcelExtension() {
        String filename = excelFile.getName();
        String extension = FilenameUtils.getExtension(filename);
        Assertions.assertEquals(true, ExcelUtility.isValidExcelExtension(extension));
    }

    @Test
    void getCellName() {
        int row = 0;
        int col = 0;
        String cellName = ExcelUtility.getCellName(row, col);
        Assertions.assertEquals("A1", cellName);
    }
}