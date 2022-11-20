package tools;

import enums.Extension;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
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
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz) throws IllegalAccessException, IOException {
        return convertObjectsToExcelFile(objects, clazz, "./", clazz.getSimpleName(), Extension.XLSX, true);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, String filename) throws IllegalAccessException, IOException {
        return convertObjectsToExcelFile(objects, clazz, "./", filename, Extension.XLSX, true);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, String path, String filename) throws IllegalAccessException, IOException {
        return convertObjectsToExcelFile(objects, clazz, path, filename, Extension.XLSX, true);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, String path, String filename, Boolean writeHeader) throws IllegalAccessException, IOException {
        return convertObjectsToExcelFile(objects, clazz, path, filename, Extension.XLSX, writeHeader);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, String path, String filename, Extension extension) throws IllegalAccessException, IOException {
        return convertObjectsToExcelFile(objects, clazz, path, filename, extension, true);
    }

    @Override
    public File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String path, String filename, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException {

        /* Create workbook and sheet */
        Workbook workbook = this.createWorkbook(extension);
        Sheet sheet = workbook.createSheet(clazz.getSimpleName());

        Field[] fields = clazz.getDeclaredFields();
        this.setFieldsAccessible(fields);
        int cRow = 0;

        /* Write header */
        if(writeHeader) {
            this.writeExcelHeader(sheet, fields, cRow++);
        }

        /* Write body */
        for (Object object : objects) {
            this.writeExcelBody(sheet, fields, object, cRow++);
        }

        /* Write file */
        String pathname = this.getPathname(path, filename, extension);
        File file = new File(pathname);
        FileOutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);

        /* Close file */
        outputStream.close();
        workbook.close();

        return file;
    }

    private Workbook createWorkbook(Extension extension) {
        Workbook workbook = null;
        switch (extension) {
            case XLS -> workbook = new HSSFWorkbook();
            case XLSX -> workbook = new XSSFWorkbook();
        }
        return workbook;
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

    private String getPathname(String path, String filename, Extension extension) {
        path = path.replaceAll("\\\\", "/");
        if(path.charAt(path.length() - 1) != '/') {
            path += '/';
        }

        return path + filename + '.' + extension.getExt();
    }
}
