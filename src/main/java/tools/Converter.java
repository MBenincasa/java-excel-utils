package tools;

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

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

/**
 * {@code Converter} is the static class with implementations of conversion methods
 * @author Mirko Benincasa
 * @since 0.2.0
 */
public class Converter {

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. The default filename is the class name. By default, the extension that is selected is XLSX while the header is added if not specified
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    public static File objectsToExcel(List<?> objects, Class<?> clazz) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), Extension.XLSX, true);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. By default, the extension that is selected is XLSX while the header is added if not specified
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param filename The name of the output file without the extension
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String filename) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, Extension.XLSX, true);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * By default the extension that is selected is XLSX while the header is added if not specified
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param path The destination path of the output file
     * @param filename The name of the output file without the extension
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, path, filename, Extension.XLSX, true);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * By default the extension that is selected is XLSX
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param path The destination path of the output file
     * @param filename The name of the output file without the extension
     * @param writeHeader If {@code true} it will write the header to the first line
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, path, filename, Extension.XLSX, writeHeader);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. The default filename is the class name. By default, the extension that is selected is XLSX
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param writeHeader If {@code true} it will write the header to the first line
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    public static File objectsToExcel(List<?> objects, Class<?> clazz, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), Extension.XLSX, writeHeader);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. By default, the extension that is selected is XLSX
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param filename The name of the output file without the extension
     * @param writeHeader If {@code true} it will write the header to the first line
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String filename, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, Extension.XLSX, writeHeader);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * By default, the header is added
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param path The destination path of the output file
     * @param filename The name of the output file without the extension
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, Extension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, path, filename, extension, true);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. The default filename is the class name. By default, the header is added if not specified
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    public static File objectsToExcel(List<?> objects, Class<?> clazz, Extension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), extension, true);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. The default filename is the class name
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @param writeHeader If {@code true} it will write the header to the first line
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    public static File objectsToExcel(List<?> objects, Class<?> clazz, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), extension, writeHeader);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. By default, the header is added if not specified
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param filename The name of the output file without the extension
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String filename, Extension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, extension, true);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param filename The name of the output file without the extension
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @param writeHeader If {@code true} it will write the header to the first line
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String filename, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, extension, writeHeader);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param path The destination path of the output file
     * @param filename The name of the output file without the extension
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @param writeHeader If {@code true} it will write the header to the first line
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException {
        /* Check extension*/
        if(!extension.isExcelExtension())
            throw new ExtensionNotValidException("Select an extension for an Excel file");

        /* Open file */
        String pathname = getPathname(path, filename, extension);
        File file = new File(pathname);

        if (file.exists()) {
            throw new FileAlreadyExistsException("There is already a file with this pathname: " + file.getAbsolutePath());
        }

        /* Create workbook and sheet */
        Workbook workbook = WorkbookUtility.create(extension);
        objectsToExistingExcel(workbook, objects, clazz, writeHeader);

        /* Write file */
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);

        /* Close file */
        WorkbookUtility.close(workbook, fileOutputStream);

        return file;
    }

    /**
     * This method allows you to convert objects into a Sheet of a Workbook that already exists.<p>
     * Note: This method does not call the "write" method of the workbook.<p>
     * By default, the header is added if not specified
     * @param workbook The {@code Workbook} to update
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     */
    public static void objectsToExistingExcel(Workbook workbook, List<?> objects, Class<?> clazz) throws IllegalAccessException {
        objectsToExistingExcel(workbook, objects, clazz, true);
    }

    /**
     * This method allows you to convert objects into a Sheet of a Workbook that already exists.<p>
     * Note: This method does not call the "write" method of the workbook.
     * @param workbook The {@code Workbook} to update
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param writeHeader If {@code true} it will write the header to the first line
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     */
    public static void objectsToExistingExcel(Workbook workbook, List<?> objects, Class<?> clazz, Boolean writeHeader) throws IllegalAccessException {
        /* Create sheet */
        Sheet sheet = SheetUtility.create(workbook, clazz.getSimpleName());

        Field[] fields = clazz.getDeclaredFields();
        setFieldsAccessible(fields);
        int cRow = 0;

        /* Write header */
        if (writeHeader) {
            CellStyle headerCellStyle = createHeaderCellStyle(workbook, clazz);
            writeExcelHeader(sheet, fields, cRow++, headerCellStyle);
        }

        /* Write body */
        for (Object object : objects) {
            CellStyle bodyCellStyle = createBodyStyle(workbook, clazz);
            writeExcelBody(workbook, sheet, fields, object, cRow++, bodyCellStyle, clazz);
        }
    }

    /**
     * Convert an Excel file into a list of objects<p>
     * Note: The type of the elements of the return objects must coincide with the type of {@code clazz}<p>
     * By default, the first sheet is chosen
     * @param file The input Excel file that will be converted into a list of objects
     * @param clazz The class of the list elements
     * @return A list of objects that contains as many objects as there are lines in the input file
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws InstantiationException If an error occurs while instantiating a new object
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws InvocationTargetException If an error occurs while instantiating a new object or setting a field
     * @throws NoSuchMethodException If the setting method or empty constructor of the object is not found
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws HeaderNotPresentException If the first row is empty and does not contain the header
     */
    public static List<?> excelToObjects(File file, Class<?> clazz) throws ExtensionNotValidException, IOException, OpenWorkbookException, InstantiationException, IllegalAccessException, InvocationTargetException, NoSuchMethodException, SheetNotFoundException, HeaderNotPresentException {
        return excelToObjects(file, clazz, null);
    }

    /**
     * Convert an Excel file into a list of objects<p>
     * Note: The type of the elements of the return objects must coincide with the type of {@code clazz}
     * @param file The input Excel file that will be converted into a list of objects
     * @param clazz The class of the list elements
     * @param sheetName The name of the sheet to open
     * @return A list of objects that contains as many objects as there are lines in the input file
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws InstantiationException If an error occurs while instantiating a new object
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws InvocationTargetException If an error occurs while instantiating a new object or setting a field
     * @throws NoSuchMethodException If the setting method or empty constructor of the object is not found
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws HeaderNotPresentException If the first row is empty and does not contain the header
     */
    public static List<?> excelToObjects(File file, Class<?> clazz, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, InvocationTargetException, IllegalAccessException, NoSuchMethodException, InstantiationException, SheetNotFoundException, HeaderNotPresentException {
        /* Open file excel */
        Workbook workbook = WorkbookUtility.open(file);
        Sheet sheet = (sheetName == null || sheetName.isEmpty())
                ? SheetUtility.get(workbook)
                : SheetUtility.get(workbook, sheetName);

        /* Retrieving header names */
        Field[] fields = clazz.getDeclaredFields();
        setFieldsAccessible(fields);
        Map<Integer, String> headerMap = getHeaderNames(sheet, fields);

        /* Converting cells to objects */
        List<Object> resultList = new ArrayList<>();
        for (Row row : sheet) {
            if (row == null || row.getRowNum() == 0) {
                continue;
            }

            Object obj = convertCellValuesToObject(clazz, row, fields, headerMap);
            resultList.add(obj);
        }

        /* Close file */
        WorkbookUtility.close(workbook);

        return resultList;
    }

    /**
     * Convert an Excel file into a CSV file<p>
     * The default path is that of the temporary folder. By default, the first sheet is chosen and the filename will be the same as the input file if not specified
     * @param fileInput The input Excel file that will be converted into a CSV file
     * @return A CSV file that contains the same lines as the Excel file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     */
    public static File excelToCsv(File fileInput) throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        return excelToCsv(fileInput, System.getProperty("java.io.tmpdir"), fileInput.getName().split("\\.")[0].trim(), null);
    }

    /**
     * Convert an Excel file into a CSV file<p>
     * The default path is that of the temporary folder. By default, the first sheet is chosen and the filename will be the same as the input file if not specified
     * @param fileInput The input Excel file that will be converted into a CSV file
     * @param sheetName The name of the sheet to open
     * @return A CSV file that contains the same lines as the Excel file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     */
    public static File excelToCsv(File fileInput, String sheetName) throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        return excelToCsv(fileInput, System.getProperty("java.io.tmpdir"), fileInput.getName().split("\\.")[0].trim(), sheetName);
    }

    /**
     * Convert an Excel file into a CSV file<p>
     * By default, the first sheet is chosen
     * @param fileInput The input Excel file that will be converted into a CSV file
     * @param path The destination path of the output file
     * @param filename The name of the output file without the extension
     * @return A CSV file that contains the same lines as the Excel file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     */
    public static File excelToCsv(File fileInput, String path, String filename) throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        return excelToCsv(fileInput, path, filename, null);
    }

    /**
     * Convert an Excel file into a CSV file
     * @param fileInput The input Excel file that will be converted into a CSV file
     * @param path The destination path of the output file
     * @param filename The name of the output file without the extension
     * @param sheetName The name of the sheet to open
     * @return A CSV file that contains the same lines as the Excel file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     */
    public static File excelToCsv(File fileInput, String path, String filename, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException, FileAlreadyExistsException {
        /* Open file excel */
        Workbook workbook = WorkbookUtility.open(fileInput);
        Sheet sheet = (sheetName == null || sheetName.isEmpty())
                ? SheetUtility.get(workbook)
                : SheetUtility.get(workbook, sheetName);

        /* Create output file */
        String pathname = getPathname(path, filename, Extension.CSV.getExt());
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
        WorkbookUtility.close(workbook, csvWriter);

        return csvFile;
    }

    /**
     * Convert a CSV file into an Excel file<p>
     * The default path is that of the temporary folder. By default, the filename will be the same as the input file if not specified and the extension is XLSX
     * @param fileInput The input CSV file that will be converted into an Excel file
     * @return An Excel file that contains the same lines as the CSV file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     * @throws IOException If an I/O error has occurred
     */
    public static File csvToExcel(File fileInput) throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException {
        return csvToExcel(fileInput, System.getProperty("java.io.tmpdir"), fileInput.getName().split("\\.")[0].trim(), Extension.XLSX);
    }

    /**
     * Convert a CSV file into an Excel file<p>
     * The default path is that of the temporary folder. By default, the extension is XLSX
     * @param fileInput The input CSV file that will be converted into an Excel file
     * @param filename The name of the output file without the extension
     * @return An Excel file that contains the same lines as the CSV file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     * @throws IOException If an I/O error has occurred
     */
    public static File csvToExcel(File fileInput, String filename) throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException {
        return csvToExcel(fileInput, System.getProperty("java.io.tmpdir"), filename, Extension.XLSX);
    }

    /**
     * Convert a CSV file into an Excel file<p>
     * By default, the extension is XLSX
     * @param fileInput The input CSV file that will be converted into an Excel file
     * @param path The destination path of the output file
     * @param filename The name of the output file without the extension
     * @return An Excel file that contains the same lines as the CSV file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     * @throws IOException If an I/O error has occurred
     */
    public static File csvToExcel(File fileInput, String path, String filename) throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException {
        return csvToExcel(fileInput, path, filename, Extension.XLSX);
    }

    /**
     * Convert a CSV file into an Excel file
     * @param fileInput The input CSV file that will be converted into an Excel file
     * @param path The destination path of the output file
     * @param filename The name of the output file without the extension
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @return An Excel file that contains the same lines as the CSV file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     * @throws IOException If an I/O error has occurred
     */
    public static File csvToExcel(File fileInput, String path, String filename, Extension extension) throws IOException, ExtensionNotValidException, CsvValidationException, FileAlreadyExistsException {
        /* Check exension */
        String csvExt = FilenameUtils.getExtension(fileInput.getName());
        isValidCsvExtension(csvExt);

        /* Open CSV file */
        FileReader fileReader = new FileReader(fileInput);
        CSVReader csvReader = new CSVReader(fileReader);

        /* Create output file */
        String pathname = getPathname(path, filename, extension);
        File outputFile = new File(pathname);

        if (outputFile.exists()) {
            throw new FileAlreadyExistsException("There is already a file with this pathname: " + outputFile.getAbsolutePath());
        }

        /* Create workbook and sheet */
        Workbook workbook = WorkbookUtility.create(extension);
        csvToExistingExcel(workbook, csvReader);

        /* Write file */
        FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
        workbook.write(fileOutputStream);

        /* Close file */
        WorkbookUtility.close(workbook, fileOutputStream, csvReader);

        return outputFile;
    }

    /**
     * Convert the CSV file into a new sheet of an existing Workbook.<p>
     * Note: This method does not call the "write" method of the workbook.
     * @param workbook The {@code Workbook} to update
     * @param fileInput The input CSV file that will be converted into an Excel file
     * @throws IOException If an I/O error has occurred
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     */
    public static void csvToExistingExcel(Workbook workbook, File fileInput) throws IOException, CsvValidationException, ExtensionNotValidException {
        /* Check exension */
        String csvExt = FilenameUtils.getExtension(fileInput.getName());
        isValidCsvExtension(csvExt);

        /* Open CSV file */
        FileReader fileReader = new FileReader(fileInput);
        CSVReader csvReader = new CSVReader(fileReader);
        csvToExistingExcel(workbook, csvReader);

        /* Close CSV reader */
        csvReader.close();
    }

    /**
     * Writes the data present in the CSVReader to a new sheet of an existing Workbook.<p>
     * Note: This method does not call the "write" method of the workbook.
     * @param workbook The {@code Workbook} to update
     * @param csvReader The {@code CSVReader} of the CSV input file
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws IOException If an I/O error has occurred
     */
    public static void csvToExistingExcel(Workbook workbook, CSVReader csvReader) throws CsvValidationException, IOException {
        Sheet sheet = SheetUtility.create(workbook);

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
    }

    private static void isValidCsvExtension(String extension) throws ExtensionNotValidException {
        if (!extension.equalsIgnoreCase(Extension.CSV.getExt()))
            throw new ExtensionNotValidException("Pass a file with the CSV extension");
    }

    private static Map<Integer, String> getHeaderNames(Sheet sheet, Field[] fields) throws HeaderNotPresentException {
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

    private static Object convertCellValuesToObject(Class<?> clazz, Row row, Field[] fields, Map<Integer, String> headerMap) throws InvocationTargetException, IllegalAccessException, NoSuchMethodException, InstantiationException {
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

    private static void setFieldsAccessible(Field[] fields) {
        for (Field field : fields) {
            field.setAccessible(true);
        }
    }

    private static void writeExcelHeader(Sheet sheet, Field[] fields, int cRow, CellStyle cellStyle) {
        Row headerRow = sheet.createRow(cRow);
        for (int i = 0; i < fields.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellStyle(cellStyle);
            ExcelField excelField = fields[i].getAnnotation(ExcelField.class);
            cell.setCellValue(excelField != null ? excelField.name() : fields[i].getName());
        }
    }

    private static CellStyle createHeaderCellStyle(Workbook workbook, Class<?> clazz) {
        CellStyle cellStyle = workbook.createCellStyle();
        ExcelHeaderStyle excelHeaderStyle = clazz.getAnnotation(ExcelHeaderStyle.class);
        if (excelHeaderStyle == null) {
            return cellStyle;
        }
        return createCellStyle(cellStyle, excelHeaderStyle.cellColor(), excelHeaderStyle.horizontal(), excelHeaderStyle.vertical());
    }

    private static void writeExcelBody(Workbook workbook, Sheet sheet, Field[] fields, Object object, int cRow, CellStyle cellStyle, Class<?> clazz) throws IllegalAccessException {
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

    private static CellStyle createBodyStyle(Workbook workbook, Class<?> clazz) {
        CellStyle cellStyle = workbook.createCellStyle();
        ExcelBodyStyle excelBodyStyle = clazz.getAnnotation(ExcelBodyStyle.class);
        if (excelBodyStyle == null) {
            return cellStyle;
        }
        return createCellStyle(cellStyle, excelBodyStyle.cellColor(), excelBodyStyle.horizontal(), excelBodyStyle.vertical());
    }

    private static CellStyle createCellStyle(CellStyle cellStyle, IndexedColors indexedColors, HorizontalAlignment horizontal, VerticalAlignment vertical) {
        cellStyle.setFillForegroundColor(indexedColors.getIndex());
        cellStyle.setFillPattern(FillPatternType.BIG_SPOTS);
        cellStyle.setAlignment(horizontal);
        cellStyle.setVerticalAlignment(vertical);
        cellStyle.setBorderBottom(BorderStyle.MEDIUM);

        return cellStyle;
    }

    private static CellStyle cloneStyle(Workbook workbook, CellStyle cellStyle) {
        CellStyle newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(cellStyle);
        return newStyle;
    }

    private static void setAutoSizeColumn(Sheet sheet, Field[] fields, Class<?> clazz) {
        ExcelHeaderStyle excelHeaderStyle = clazz.getAnnotation(ExcelHeaderStyle.class);
        if (excelHeaderStyle != null && excelHeaderStyle.autoSize()) {
            for (int i = 0; i < fields.length; i++) {
                sheet.autoSizeColumn(i);
            }
        }
    }

    private static String getPathname(String path, String filename, Extension extension) {
        return getPathname(path, filename, extension.getExt());
    }

    private static String getPathname(String path, String filename, String extension) {
        path = path.replaceAll("\\\\", "/");
        if (path.charAt(path.length() - 1) != '/') {
            path += '/';
        }

        return path + filename + '.' + extension;
    }
}
