package tools.implementations;

import annotations.ExcelBodyStyle;
import annotations.ExcelField;
import annotations.ExcelHeaderStyle;
import enums.ExcelExtension;
import exceptions.ExtensionNotValidException;
import exceptions.FileAlreadyExistsException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.*;
import tools.interfaces.ExcelConverter;
import tools.interfaces.ExcelSheetUtils;
import tools.interfaces.ExcelUtils;
import tools.interfaces.ExcelWorkbookUtils;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

public class ExcelConverterImpl implements ExcelConverter {

    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), ExcelExtension.XLSX, true);
    }

    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String filename) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, ExcelExtension.XLSX, true);
    }

    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return objectsToExcel(objects, clazz, path, filename, ExcelExtension.XLSX, true);
    }

    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return objectsToExcel(objects, clazz, path, filename, ExcelExtension.XLSX, writeHeader);
    }

    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), ExcelExtension.XLSX, writeHeader);
    }

    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String filename, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, ExcelExtension.XLSX, writeHeader);
    }

    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, ExcelExtension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return objectsToExcel(objects, clazz, path, filename, extension, true);
    }

    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, ExcelExtension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), extension, true);
    }

    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, ExcelExtension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), extension, writeHeader);
    }

    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String filename, ExcelExtension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, extension, true);
    }

    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String filename, ExcelExtension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, extension, writeHeader);
    }

    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, ExcelExtension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException {

        /* Open file */
        String pathname = this.getPathname(path, filename, extension);
        File file = new File(pathname);

        if(file.exists()) {
            throw new FileAlreadyExistsException("There is already a file with this pathname: " + file.getAbsolutePath());
        }

        /* Create workbook and sheet */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        Workbook workbook = excelWorkbookUtils.create(extension);
        ExcelSheetUtils excelSheetUtils = new ExcelSheetUtilsImpl();
        Sheet sheet = excelSheetUtils.create(workbook, clazz.getSimpleName());

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
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);

        /* Close file */
        excelWorkbookUtils.close(workbook, fileOutputStream);

        return file;
    }

    @Override
    public List<?> excelToObjects(File file, Class<?> clazz) throws ExtensionNotValidException, IOException, OpenWorkbookException, InstantiationException, IllegalAccessException, InvocationTargetException, NoSuchMethodException, SheetNotFoundException {
        return excelToObjects(file, clazz, null);
    }

    @Override
    public List<?> excelToObjects(File file, Class<?> clazz, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, InvocationTargetException, IllegalAccessException, NoSuchMethodException, InstantiationException, SheetNotFoundException {

        /* Check extension */
        String extension = checkExtension(file.getName());

        /* Open file excel */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = excelWorkbookUtils.open(fileInputStream, extension);
        ExcelSheetUtils excelSheetUtils = new ExcelSheetUtilsImpl();
        Sheet sheet = (sheetName == null || sheetName.isEmpty())
                ? excelSheetUtils.open(workbook)
                : excelSheetUtils.open(workbook, sheetName);

        /* Retrieving header names */
        Field[] fields = clazz.getDeclaredFields();
        this.setFieldsAccessible(fields);
        Map<Integer, String> headerMap = getHeaderNames(sheet, fields);

        /* Converting cells to objects */
        List<Object> resultList = new ArrayList<>();
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            Object obj = convertCellValuesToObject(clazz, row, fields, headerMap);
            resultList.add(obj);
        }

        /* Close file */
        excelWorkbookUtils.close(workbook, fileInputStream);

        return resultList;
    }

    private Map<Integer, String> getHeaderNames(Sheet sheet, Field[] fields) {
        Map<String, String> fieldNames = new HashMap<>();
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            fieldNames.put(excelField == null ? field.getName() : excelField.name(), field.getName());
        }

        Row headerRow = sheet.getRow(0);
        Map<Integer, String> headerMap = new TreeMap<>();
        for (int i = 0; i < headerRow.getPhysicalNumberOfCells(); i++) {
            Cell cell = headerRow.getCell(i);
            if (fieldNames.containsKey(cell.getStringCellValue())) {
                headerMap.put(i, fieldNames.get(cell.getStringCellValue()));
            }
        }

        return headerMap;
    }

    private Object convertCellValuesToObject(Class<?> clazz, Row row, Field[] fields, Map<Integer, String> headerMap) throws InvocationTargetException, IllegalAccessException, NoSuchMethodException, InstantiationException {
        Object obj = clazz.getDeclaredConstructor().newInstance();
        for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
            String headerName = headerMap.get(j);
            Cell cell = row.getCell(j);

            switch (cell.getCellType()) {
                case NUMERIC -> {
                    Optional<Field> fieldOptional = Arrays.stream(fields).filter(f -> f.getName().equalsIgnoreCase(headerName)).findFirst();
                    if (fieldOptional.isEmpty()) {
                        throw new RuntimeException();
                    }
                    Field field = fieldOptional.get();

                    if (Integer.class.equals(field.getType())) {
                        PropertyUtils.setSimpleProperty(obj, headerName, (int) cell.getNumericCellValue());
                    } else if (Double.class.equals(field.getType())) {
                        PropertyUtils.setSimpleProperty(obj, headerName, cell.getNumericCellValue());
                    } else if (Long.class.equals(field.getType())) {
                        PropertyUtils.setSimpleProperty(obj, headerName, (long) cell.getNumericCellValue());
                    } else if (Date.class.equals(field.getType())) {
                        PropertyUtils.setSimpleProperty(obj, headerName, cell.getDateCellValue());
                    } else if (LocalDateTime.class.equals(field.getType())) {
                        PropertyUtils.setSimpleProperty(obj, headerName, cell.getLocalDateTimeCellValue());
                    } else if (LocalDate.class.equals(field.getType())) {
                        PropertyUtils.setSimpleProperty(obj, headerName, cell.getLocalDateTimeCellValue().toLocalDate());
                    }

                }
                case BOOLEAN -> PropertyUtils.setSimpleProperty(obj, headerName, cell.getBooleanCellValue());
                default -> PropertyUtils.setSimpleProperty(obj, headerName, cell.getStringCellValue());
            }
        }
        return obj;
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

    private CellStyle createHeaderCellStyle(Workbook workbook, Class<?> clazz) {
        CellStyle cellStyle = workbook.createCellStyle();
        ExcelHeaderStyle excelHeaderStyle = clazz.getAnnotation(ExcelHeaderStyle.class);
        if (excelHeaderStyle == null) {
            return cellStyle;
        }
        return createCellStyle(cellStyle, excelHeaderStyle.cellColor(), excelHeaderStyle.horizontal(), excelHeaderStyle.vertical());
    }

    private void writeExcelBody(Workbook workbook, Sheet sheet, Field[] fields, Object object, int cRow, CellStyle cellStyle, Class<?> clazz) throws IllegalAccessException {
        Row row = sheet.createRow(cRow);
        for (int i = 0; i < fields.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(cellStyle);

            if (fields[i].get(object) instanceof Integer || fields[i].get(object) instanceof Long) {
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
            } else if (fields[i].get(object) instanceof LocalDate) {
                CellStyle newStyle = cloneStyle(workbook, cellStyle);
                newStyle.setDataFormat((short) 14);
                cell.setCellStyle(newStyle);
                cell.setCellValue((LocalDate) fields[i].get(object));
            } else if (fields[i].get(object) instanceof LocalDateTime) {
                CellStyle newStyle = cloneStyle(workbook, cellStyle);
                newStyle.setDataFormat((short) 22);
                cell.setCellStyle(newStyle);
                cell.setCellValue((LocalDateTime) fields[i].get(object));
            } else if (fields[i].get(object) instanceof Boolean) {
                cell.setCellValue((Boolean) fields[i].get(object));
            } else {
                cell.setCellValue(String.valueOf(fields[i].get(object)));
            }
        }

        /* Set auto-size columns */
        setAutoSizeColumn(sheet, fields, clazz);
    }

    private CellStyle createBodyStyle(Workbook workbook, Class<?> clazz) {
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

    private void setAutoSizeColumn(Sheet sheet, Field[] fields, Class<?> clazz) {
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

    private String checkExtension(String filename) throws ExtensionNotValidException {
        String extension = FilenameUtils.getExtension(filename);
        ExcelUtils excelUtils = new ExcelUtilsImpl();

        if(!excelUtils.isValidExcelExtension(extension)) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }
        return extension;
    }
}
