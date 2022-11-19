package tools.converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.List;

public class ExcelConverterImpl implements ExcelConverter {

    @Override
    public File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String pathname, Boolean writeHeader) throws IllegalAccessException, IOException {

        /* Create workbook and sheet */
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(clazz.getName());

        Field[] fields = clazz.getDeclaredFields();
        this.setFieldsAccessible(fields);
        int cRow = 0;

        /* Write header */
        if(writeHeader) {
            writeExcelHeader(sheet, fields, cRow++);
        }

        /* Write body */
        for (Object object : objects) {
            writeExcelBody(sheet, fields, object, cRow++);
        }

        /* Write file */
        File file = new File(pathname);
        FileOutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);

        /* Close file */
        outputStream.close();
        workbook.close();

        return file;
    }

    private void setFieldsAccessible(Field[] fields) {
        for (Field field : fields) {
            field.setAccessible(true);
        }
    }

    private void writeExcelHeader(Sheet sheet, Field[] fields, int cRow) {
        Row headerRow = sheet.createRow(cRow);
        for (int i = 0; i < fields.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(fields[i].getName());
        }
    }

    private void writeExcelBody(Sheet sheet, Field[] fields, Object object, int cRow) throws IllegalAccessException {
        Row row = sheet.createRow(cRow);
        for (int i = 0; i < fields.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(String.valueOf(fields[i].get(object)));
        }
    }
}
