package tools.implementations;

import annotations.ExcelBodyStyle;
import annotations.ExcelField;
import annotations.ExcelHeaderStyle;
import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import com.opencsv.exceptions.CsvValidationException;
import enums.Extension;
import exceptions.*;
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

/**
 * {@code ExcelConverterImpl} is the standard implementation class of {@code ExcelConverter}
 * @deprecated since version 0.2.0. View here {@link tools.Converter}
 * @see tools.Converter
 * @author Mirko Benincasa
 * @since 0.1.0
 */
@Deprecated
public class ExcelConverterImpl implements ExcelConverter {

    /**
     * {@inheritDoc}<p>
     * The default path is that of the temporary folder. The default filename is the class name. By default the extension that is selected is XLSX while the header is added if not specified
     * @param objects {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @return {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     */
    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), Extension.XLSX, true);
    }

    /**
     * {@inheritDoc}<p>
     * The default path is that of the temporary folder. By default the extension that is selected is XLSX while the header is added if not specified
     * @param objects {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @param filename {@inheritDoc}
     * @return {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     */
    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String filename) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, Extension.XLSX, true);
    }

    /**
     * {@inheritDoc}<p>
     * By default the extension that is selected is XLSX while the header is added if not specified
     * @param objects {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @param path {@inheritDoc}
     * @param filename {@inheritDoc}
     * @return {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     */
    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, path, filename, Extension.XLSX, true);
    }

    /**
     * {@inheritDoc}<p>
     * By default the extension that is selected is XLSX
     * @param objects {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @param path {@inheritDoc}
     * @param filename {@inheritDoc}
     * @param writeHeader {@inheritDoc}
     * @return {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     */
    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, path, filename, Extension.XLSX, writeHeader);
    }

    /**
     * {@inheritDoc}<p>
     * The default path is that of the temporary folder. The default filename is the class name. By default the extension that is selected is XLSX
     * @param objects {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @param writeHeader {@inheritDoc}
     * @return {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     */
    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), Extension.XLSX, writeHeader);
    }

    /**
     * {@inheritDoc}<p>
     * The default path is that of the temporary folder. By default the extension that is selected is XLSX
     * @param objects {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @param filename {@inheritDoc}
     * @param writeHeader {@inheritDoc}
     * @return {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     */
    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String filename, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, Extension.XLSX, writeHeader);
    }

    /**
     * {@inheritDoc}<p>
     * By default, the header is added
     * @param objects {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @param path {@inheritDoc}
     * @param filename {@inheritDoc}
     * @param extension {@inheritDoc}
     * @return {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     */
    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, Extension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, path, filename, extension, true);
    }

    /**
     * {@inheritDoc}<p>
     * The default path is that of the temporary folder. The default filename is the class name. By default, the header is added if not specified
     * @param objects {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @param extension {@inheritDoc}
     * @return {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     */
    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, Extension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), extension, true);
    }

    /**
     * {@inheritDoc}<p>
     * The default path is that of the temporary folder. The default filename is the class name
     * @param objects {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @param extension {@inheritDoc}
     * @param writeHeader {@inheritDoc}
     * @return {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     */
    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), extension, writeHeader);
    }

    /**
     * {@inheritDoc}<p>
     * The default path is that of the temporary folder. By default, the header is added if not specified
     * @param objects {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @param filename {@inheritDoc}
     * @param extension {@inheritDoc}
     * @return {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     */
    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String filename, Extension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, extension, true);
    }

    /**
     * {@inheritDoc}<p>
     * The default path is that of the temporary folder.
     * @param objects {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @param filename {@inheritDoc}
     * @param extension {@inheritDoc}
     * @param writeHeader {@inheritDoc}
     * @return {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     */
    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String filename, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, extension, writeHeader);
    }

    /**
     * {@inheritDoc}
     * @param objects {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @param path {@inheritDoc}
     * @param filename {@inheritDoc}
     * @param extension {@inheritDoc}
     * @param writeHeader {@inheritDoc}
     * @return {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     */
    @Override
    public File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        /* Check extension*/
        if(!extension.isExcelExtension())
            throw new ExtensionNotValidException("Select an extension for an Excel file");

        /* Open file */
        String pathname = this.getPathname(path, filename, extension);
        File file = new File(pathname);

        if (file.exists()) {
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
        if (writeHeader) {
            CellStyle headerCellStyle = this.createHeaderCellStyle(workbook, clazz);
            this.writeExcelHeader(sheet, fields, cRow++, headerCellStyle);
        }

        /* Write body */
        for (Object object : objects) {
            CellStyle bodyCellStyle = this.createBodyStyle(workbook, clazz);
            this.writeExcelBody(workbook, sheet, fields, object, cRow++, bodyCellStyle, clazz);
        }

        /* Write file */
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);

        /* Close file */
        excelWorkbookUtils.close(workbook, fileOutputStream);

        return file;
    }

    /**
     * {@inheritDoc}<p>
     * By default, the first sheet is chosen
     * @param file {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @return {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws OpenWorkbookException {@inheritDoc}
     * @throws InstantiationException {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws InvocationTargetException {@inheritDoc}
     * @throws NoSuchMethodException {@inheritDoc}
     * @throws SheetNotFoundException {@inheritDoc}
     * @throws HeaderNotPresentException {@inheritDoc}
     */
    @Override
    public List<?> excelToObjects(File file, Class<?> clazz) throws ExtensionNotValidException, IOException, OpenWorkbookException, InstantiationException, IllegalAccessException, InvocationTargetException, NoSuchMethodException, SheetNotFoundException, HeaderNotPresentException {
        return excelToObjects(file, clazz, null);
    }

    /**
     * {@inheritDoc}
     * @param file {@inheritDoc}
     * @param clazz {@inheritDoc}
     * @param sheetName {@inheritDoc}
     * @return {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws OpenWorkbookException {@inheritDoc}
     * @throws InstantiationException {@inheritDoc}
     * @throws IllegalAccessException {@inheritDoc}
     * @throws InvocationTargetException {@inheritDoc}
     * @throws NoSuchMethodException {@inheritDoc}
     * @throws SheetNotFoundException {@inheritDoc}
     * @throws HeaderNotPresentException {@inheritDoc}
     */
    @Override
    public List<?> excelToObjects(File file, Class<?> clazz, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, InvocationTargetException, IllegalAccessException, NoSuchMethodException, InstantiationException, SheetNotFoundException, HeaderNotPresentException {
        /* Check extension */
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        String extension = excelUtils.checkExcelExtension(file.getName());

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
        Map<Integer, String> headerMap = this.getHeaderNames(sheet, fields);

        /* Converting cells to objects */
        List<Object> resultList = new ArrayList<>();
        for (Row row : sheet) {
            if (row == null || row.getRowNum() == 0) {
                continue;
            }

            Object obj = this.convertCellValuesToObject(clazz, row, fields, headerMap);
            resultList.add(obj);
        }

        /* Close file */
        excelWorkbookUtils.close(workbook, fileInputStream);

        return resultList;
    }

    /**
     * {@inheritDoc}<p>
     * The default path is that of the temporary folder. By default, the first sheet is chosen and the filename will be the same as the input file if not specified
     * @param fileInput {@inheritDoc}
     * @return {@inheritDoc}
     * @throws FileAlreadyExistsException @inheritDoc}
     * @throws OpenWorkbookException @inheritDoc}
     * @throws SheetNotFoundException @inheritDoc}
     * @throws ExtensionNotValidException @inheritDoc}
     * @throws IOException {@inheritDoc}
     */
    @Override
    public File excelToCsv(File fileInput) throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        return excelToCsv(fileInput, System.getProperty("java.io.tmpdir"), fileInput.getName().split("\\.")[0].trim(), null);
    }

    /**
     * {@inheritDoc}<p>
     * The default path is that of the temporary folder. By default, the first sheet is chosen and the filename will be the same as the input file if not specified
     * @param fileInput {@inheritDoc}
     * @param sheetName {@inheritDoc}
     * @return {@inheritDoc}
     * @throws FileAlreadyExistsException @inheritDoc}
     * @throws OpenWorkbookException @inheritDoc}
     * @throws SheetNotFoundException @inheritDoc}
     * @throws ExtensionNotValidException @inheritDoc}
     * @throws IOException {@inheritDoc}
     */
    @Override
    public File excelToCsv(File fileInput, String sheetName) throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        return excelToCsv(fileInput, System.getProperty("java.io.tmpdir"), fileInput.getName().split("\\.")[0].trim(), sheetName);
    }

    /**
     * {@inheritDoc}<p>
     * By default, the first sheet is chosen
     * @param fileInput {@inheritDoc}
     * @param path {@inheritDoc}
     * @param filename {@inheritDoc}
     * @return {@inheritDoc}
     * @throws FileAlreadyExistsException @inheritDoc}
     * @throws OpenWorkbookException @inheritDoc}
     * @throws SheetNotFoundException @inheritDoc}
     * @throws ExtensionNotValidException @inheritDoc}
     * @throws IOException {@inheritDoc}
     */
    @Override
    public File excelToCsv(File fileInput, String path, String filename) throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        return excelToCsv(fileInput, path, filename, null);
    }

    /**
     * {@inheritDoc}
     * @param fileInput {@inheritDoc}
     * @param path {@inheritDoc}
     * @param filename {@inheritDoc}
     * @param sheetName {@inheritDoc}
     * @return {@inheritDoc}
     * @throws FileAlreadyExistsException @inheritDoc}
     * @throws OpenWorkbookException @inheritDoc}
     * @throws SheetNotFoundException @inheritDoc}
     * @throws ExtensionNotValidException @inheritDoc}
     * @throws IOException {@inheritDoc}
     */
    @Override
    public File excelToCsv(File fileInput, String path, String filename, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException, FileAlreadyExistsException {
        /* Check extension */
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        String extension = excelUtils.checkExcelExtension(fileInput.getName());

        /* Open file excel */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        FileInputStream fileInputStream = new FileInputStream(fileInput);
        Workbook workbook = excelWorkbookUtils.open(fileInputStream, extension);
        ExcelSheetUtils excelSheetUtils = new ExcelSheetUtilsImpl();
        Sheet sheet = (sheetName == null || sheetName.isEmpty())
                ? excelSheetUtils.open(workbook)
                : excelSheetUtils.open(workbook, sheetName);

        /* Create output file */
        String pathname = this.getPathname(path, filename, Extension.CSV.getExt());
        File csvFile = new File(pathname);

        if (csvFile.exists()) {
            throw new FileAlreadyExistsException("There is already a file with this pathname: " + csvFile.getAbsolutePath());
        }

        /* Write output file */
        FileWriter fileWriter = new FileWriter(csvFile);
        CSVWriter csvWriter = new CSVWriter(fileWriter);

        DataFormatter formatter = new DataFormatter(true);
        for (Row row : sheet) {
            List<String> data = new LinkedList<>();
            for (int i = 0; i < row.getLastCellNum(); i++) {
                data.add(formatter.formatCellValue(row.getCell(i)));
            }
            csvWriter.writeNext(data.toArray(data.toArray(new String[0])));
        }

        /* Close file */
        excelWorkbookUtils.close(workbook, fileInputStream, csvWriter);

        return csvFile;
    }

    /**
     * {@inheritDoc}<p>
     * The default path is that of the temporary folder. By default, the filename will be the same as the input file if not specified and the extension is XLSX
     * @param fileInput {@inheritDoc}
     * @return {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws CsvValidationException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     */
    @Override
    public File csvToExcel(File fileInput) throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException {
        return csvToExcel(fileInput, System.getProperty("java.io.tmpdir"), fileInput.getName().split("\\.")[0].trim(), Extension.XLSX);
    }

    /**
     * {@inheritDoc}<p>
     * The default path is that of the temporary folder. By default, the extension is XLSX
     * @param fileInput {@inheritDoc}
     * @param filename {@inheritDoc}
     * @return {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws CsvValidationException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     */
    @Override
    public File csvToExcel(File fileInput, String filename) throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException {
        return csvToExcel(fileInput, System.getProperty("java.io.tmpdir"), filename, Extension.XLSX);
    }

    /**
     * {@inheritDoc}<p>
     * By default, the extension is XLSX
     * @param fileInput {@inheritDoc}
     * @param path {@inheritDoc}
     * @param filename {@inheritDoc}
     * @return {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws CsvValidationException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     */
    @Override
    public File csvToExcel(File fileInput, String path, String filename) throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException {
        return csvToExcel(fileInput, path, filename, Extension.XLSX);
    }

    /**
     * {@inheritDoc}
     * @param fileInput {@inheritDoc}
     * @param path {@inheritDoc}
     * @param filename {@inheritDoc}
     * @param extension {@inheritDoc}
     * @return {@inheritDoc}
     * @throws FileAlreadyExistsException {@inheritDoc}
     * @throws CsvValidationException {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     */
    @Override
    public File csvToExcel(File fileInput, String path, String filename, Extension extension) throws IOException, ExtensionNotValidException, CsvValidationException, FileAlreadyExistsException {
        /* Check exension */
        String csvExt = FilenameUtils.getExtension(fileInput.getName());
        this.isValidCsvExtension(csvExt);

        /* Open CSV file */
        FileReader fileReader = new FileReader(fileInput);
        CSVReader csvReader = new CSVReader(fileReader);

        /* Create output file */
        String pathname = this.getPathname(path, filename, extension);
        File outputFile = new File(pathname);

        if (outputFile.exists()) {
            throw new FileAlreadyExistsException("There is already a file with this pathname: " + outputFile.getAbsolutePath());
        }

        /* Create workbook and sheet */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        Workbook workbook = excelWorkbookUtils.create(extension);
        ExcelSheetUtils excelSheetUtils = new ExcelSheetUtilsImpl();
        Sheet sheet = excelSheetUtils.create(workbook);

        /* Read CSV file */
        String[] values;
        int cRow = 0;
        while ((values = csvReader.readNext()) != null) {
            Row row = sheet.createRow(cRow);
            for (int j = 0; j < values.length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(values[j]);
                sheet.autoSizeColumn(j);
            }
            cRow++;
        }

        /* Write file */
        FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
        workbook.write(fileOutputStream);

        /* Close file */
        excelWorkbookUtils.close(workbook, fileOutputStream, csvReader);

        return outputFile;
    }

    private void isValidCsvExtension(String extension) throws ExtensionNotValidException {
        if (!extension.equalsIgnoreCase(Extension.CSV.getExt()))
            throw new ExtensionNotValidException("Pass a file with the CSV extension");
    }

    private Map<Integer, String> getHeaderNames(Sheet sheet, Field[] fields) throws HeaderNotPresentException {
        Map<String, String> fieldNames = new HashMap<>();
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            fieldNames.put(excelField == null ? field.getName() : excelField.name(), field.getName());
        }

        Row headerRow = sheet.getRow(0);
        if (headerRow == null)
            throw new HeaderNotPresentException("There is no header in the first row of the sheet.");

        Map<Integer, String> headerMap = new TreeMap<>();
        for (Cell cell : headerRow) {
            if (fieldNames.containsKey(cell.getStringCellValue())) {
                headerMap.put(cell.getColumnIndex(), fieldNames.get(cell.getStringCellValue()));
            }
        }

        return headerMap;
    }

    private Object convertCellValuesToObject(Class<?> clazz, Row row, Field[] fields, Map<Integer, String> headerMap) throws InvocationTargetException, IllegalAccessException, NoSuchMethodException, InstantiationException {
        Object obj = clazz.getDeclaredConstructor().newInstance();
        for (Cell cell : row) {
            if (cell == null)
                continue;

            String headerName = headerMap.get(cell.getColumnIndex());
            if (headerName == null || headerMap.isEmpty())
                continue;

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
        return this.createCellStyle(cellStyle, excelHeaderStyle.cellColor(), excelHeaderStyle.horizontal(), excelHeaderStyle.vertical());
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
        this.setAutoSizeColumn(sheet, fields, clazz);
    }

    private CellStyle createBodyStyle(Workbook workbook, Class<?> clazz) {
        CellStyle cellStyle = workbook.createCellStyle();
        ExcelBodyStyle excelBodyStyle = clazz.getAnnotation(ExcelBodyStyle.class);
        if (excelBodyStyle == null) {
            return cellStyle;
        }
        return this.createCellStyle(cellStyle, excelBodyStyle.cellColor(), excelBodyStyle.horizontal(), excelBodyStyle.vertical());
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

    private String getPathname(String path, String filename, Extension extension) {
        return getPathname(path, filename, extension.getExt());
    }

    private String getPathname(String path, String filename, String extension) {
        path = path.replaceAll("\\\\", "/");
        if (path.charAt(path.length() - 1) != '/') {
            path += '/';
        }

        return path + filename + '.' + extension;
    }
}
