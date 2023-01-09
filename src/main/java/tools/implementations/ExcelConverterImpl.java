package tools.implementations;

import annotations.ExcelBodyStyle;
import annotations.ExcelField;
import annotations.ExcelHeaderStyle;
import enums.ExcelExtension;
import exceptions.FileAlreadyExistsException;
import org.apache.poi.ss.usermodel.*;
import tools.interfaces.ExcelConverter;
import tools.interfaces.ExcelWorkbookUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Date;
import java.util.List;

public class ExcelConverterImpl implements ExcelConverter {

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return convertObjectsToExcelFile(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), ExcelExtension.XLSX, true);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, String filename) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return convertObjectsToExcelFile(objects, clazz, System.getProperty("java.io.tmpdir"), filename, ExcelExtension.XLSX, true);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, String path, String filename) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return convertObjectsToExcelFile(objects, clazz, path, filename, ExcelExtension.XLSX, true);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, String path, String filename, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return convertObjectsToExcelFile(objects, clazz, path, filename, ExcelExtension.XLSX, writeHeader);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return convertObjectsToExcelFile(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), ExcelExtension.XLSX, writeHeader);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, String filename, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return convertObjectsToExcelFile(objects, clazz, System.getProperty("java.io.tmpdir"), filename, ExcelExtension.XLSX, writeHeader);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, String path, String filename, ExcelExtension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return convertObjectsToExcelFile(objects, clazz, path, filename, extension, true);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, ExcelExtension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return convertObjectsToExcelFile(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), extension, true);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, ExcelExtension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return convertObjectsToExcelFile(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), extension, writeHeader);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, String filename, ExcelExtension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return convertObjectsToExcelFile(objects, clazz, System.getProperty("java.io.tmpdir"), filename, extension, true);
    }

    @Override
    public File convertObjectsToExcelFile(List<?> objects, Class<?> clazz, String filename, ExcelExtension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return convertObjectsToExcelFile(objects, clazz, System.getProperty("java.io.tmpdir"), filename, extension, writeHeader);
    }

    @Override
    public File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String path, String filename, ExcelExtension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException {

        /* Open file */
        String pathname = this.getPathname(path, filename, extension);
        File file = new File(pathname);

        if(file.exists()) {
            throw new FileAlreadyExistsException("There is already a file with this pathname: " + file.getAbsolutePath());
        }

        /* Create workbook and sheet */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        Workbook workbook = excelWorkbookUtils.createWorkbook(extension);
        Sheet sheet = workbook.createSheet(clazz.getSimpleName());

        Field[] fields = clazz.getDeclaredFields();
        this.setFieldsAccessible(fields);
        int cRow = 0;

        /* Write header */
        if(writeHeader) {
            CellStyle headerCellStyle = createHeaderCellStyle(workbook, clazz);
            this.writeExcelHeader(sheet, fields, cRow++, headerCellStyle);
        }

        /* Write body */
        for (Object object : objects) {
            CellStyle bodyCellStyle = createBodyStyle(workbook, clazz);
            this.writeExcelBody(workbook, sheet, fields, object, cRow++, bodyCellStyle, clazz);
        }

        /* Write file */
        FileOutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);

        /* Close file */
        closeFile(workbook, outputStream);

        return file;
    }

    private void setFieldsAccessible(Field[] fields) {
        for (Field field : fields) {
            field.setAccessible(true);
        }
    }

    private void writeExcelHeader(Sheet sheet, Field[] fields, int cRow, CellStyle cellStyle) {
        Row headerRow = sheet.createRow(cRow);
        for (int i = 0; i < fields.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellStyle(cellStyle);
            ExcelField excelField = fields[i].getAnnotation(ExcelField.class);
            cell.setCellValue(excelField != null ? excelField.name() : fields[i].getName());
        }
    }

    private CellStyle createHeaderCellStyle(Workbook workbook, Class<? extends Object> clazz) {
        CellStyle cellStyle = workbook.createCellStyle();
        ExcelHeaderStyle excelHeaderStyle = clazz.getAnnotation(ExcelHeaderStyle.class);
        if (excelHeaderStyle == null) {
            return cellStyle;
        }
        return createCellStyle(cellStyle, excelHeaderStyle.cellColor(), excelHeaderStyle.horizontal(), excelHeaderStyle.vertical());
    }

    private void writeExcelBody(Workbook workbook, Sheet sheet, Field[] fields, Object object, int cRow, CellStyle cellStyle, Class<? extends Object> clazz) throws IllegalAccessException {
        Row row = sheet.createRow(cRow);
        for (int i = 0; i < fields.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(cellStyle);

            if (fields[i].get(object) instanceof Integer) {
                CellStyle newStyle = cloneStyle(workbook, cellStyle);
                newStyle.setDataFormat((short) 1);
                cell.setCellStyle(newStyle);
                cell.setCellValue(Integer.parseInt(String.valueOf(fields[i].get(object))));
            } else if (fields[i].get(object) instanceof Double) {
                CellStyle newStyle = cloneStyle(workbook, cellStyle);
                newStyle.setDataFormat((short) 4);
                cell.setCellStyle(newStyle);
                cell.setCellValue(Double.parseDouble(String.valueOf(fields[i].get(object))));
            } else if (fields[i].get(object) instanceof Date) {
                CellStyle newStyle = cloneStyle(workbook, cellStyle);
                newStyle.setDataFormat((short) 22);
                cell.setCellStyle(newStyle);
                cell.setCellValue((Date) fields[i].get(object));
            } else {
                cell.setCellValue(String.valueOf(fields[i].get(object)));
            }
        }

        /* Set auto-size columns */
        setAutoSizeColumn(sheet, fields, clazz);
    }

    private CellStyle createBodyStyle(Workbook workbook, Class<? extends Object> clazz) {
        CellStyle cellStyle = workbook.createCellStyle();
        ExcelBodyStyle excelBodyStyle = clazz.getAnnotation(ExcelBodyStyle.class);
        if (excelBodyStyle == null) {
            return cellStyle;
        }
        return createCellStyle(cellStyle, excelBodyStyle.cellColor(), excelBodyStyle.horizontal(), excelBodyStyle.vertical());
    }

    private CellStyle createCellStyle(CellStyle cellStyle, IndexedColors indexedColors, HorizontalAlignment horizontal, VerticalAlignment vertical) {
        cellStyle.setFillForegroundColor(indexedColors.getIndex());
        cellStyle.setFillPattern(FillPatternType.BIG_SPOTS);
        cellStyle.setAlignment(horizontal);
        cellStyle.setVerticalAlignment(vertical);
        cellStyle.setBorderBottom(BorderStyle.MEDIUM);

        return cellStyle;
    }

    private CellStyle cloneStyle(Workbook workbook, CellStyle cellStyle) {
        CellStyle newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(cellStyle);
        return newStyle;
    }

    private void setAutoSizeColumn(Sheet sheet, Field[] fields, Class<? extends Object> clazz) {
        ExcelHeaderStyle excelHeaderStyle = clazz.getAnnotation(ExcelHeaderStyle.class);
        if (excelHeaderStyle != null && excelHeaderStyle.autoSize()) {
            for (int i = 0; i < fields.length; i++) {
                sheet.autoSizeColumn(i);
            }
        }
    }

    private String getPathname(String path, String filename, ExcelExtension extension) {
        path = path.replaceAll("\\\\", "/");
        if(path.charAt(path.length() - 1) != '/') {
            path += '/';
        }

        return path + filename + '.' + extension.getExt();
    }

    private void closeFile(Workbook workbook, FileOutputStream outputStream) throws IOException {
        outputStream.close();
        workbook.close();
    }
}
