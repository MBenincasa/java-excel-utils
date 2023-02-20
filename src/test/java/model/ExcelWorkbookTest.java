package model;

import enums.Extension;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

public class ExcelWorkbookTest {

    @Test
    void close() {
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(Extension.XLSX);
        Assertions.assertDoesNotThrow(excelWorkbook::close);
    }

}