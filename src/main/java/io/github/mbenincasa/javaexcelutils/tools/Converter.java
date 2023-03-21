package io.github.mbenincasa.javaexcelutils.tools;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.github.mbenincasa.javaexcelutils.annotations.ExcelBodyStyle;
import io.github.mbenincasa.javaexcelutils.annotations.ExcelField;
import io.github.mbenincasa.javaexcelutils.annotations.ExcelHeaderStyle;
import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import com.opencsv.exceptions.CsvValidationException;
import io.github.mbenincasa.javaexcelutils.enums.Extension;
import io.github.mbenincasa.javaexcelutils.exceptions.*;
import io.github.mbenincasa.javaexcelutils.model.converter.ExcelToObject;
import io.github.mbenincasa.javaexcelutils.model.converter.JsonToExcel;
import io.github.mbenincasa.javaexcelutils.model.converter.ObjectToExcel;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelCell;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelRow;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelSheet;
import io.github.mbenincasa.javaexcelutils.model.excel.ExcelWorkbook;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

/**
 * {@code Converter} is the static class with implementations of conversion methods
 * @author Mirko Benincasa
 * @since 0.2.0
 */
public class Converter {

    private final static Logger logger = LogManager.getLogger(Converter.class);

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. The default filename is the class name. By default, the extension that is selected is XLSX while the header is added if not specified
     * @deprecated since version 0.4.0
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static File objectsToExcel(List<?> objects, Class<?> clazz) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException, SheetAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), Extension.XLSX, true);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. By default, the extension that is selected is XLSX while the header is added if not specified
     * @deprecated since version 0.4.0
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param filename The name of the output file without the extension
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String filename) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException, SheetAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, Extension.XLSX, true);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * By default the extension that is selected is XLSX while the header is added if not specified
     * @deprecated since version 0.4.0
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param path The destination path of the output file
     * @param filename The name of the output file without the extension
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException, SheetAlreadyExistsException {
        return objectsToExcel(objects, clazz, path, filename, Extension.XLSX, true);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * By default the extension that is selected is XLSX
     * @deprecated since version 0.4.0
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
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException, SheetAlreadyExistsException {
        return objectsToExcel(objects, clazz, path, filename, Extension.XLSX, writeHeader);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. The default filename is the class name. By default, the extension that is selected is XLSX
     * @deprecated since version 0.4.0
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param writeHeader If {@code true} it will write the header to the first line
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static File objectsToExcel(List<?> objects, Class<?> clazz, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException, SheetAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), Extension.XLSX, writeHeader);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. By default, the extension that is selected is XLSX
     * @deprecated since version 0.4.0
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param filename The name of the output file without the extension
     * @param writeHeader If {@code true} it will write the header to the first line
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String filename, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException, SheetAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, Extension.XLSX, writeHeader);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * By default, the header is added
     * @deprecated since version 0.4.0
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
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, Extension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException, SheetAlreadyExistsException {
        return objectsToExcel(objects, clazz, path, filename, extension, true);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. The default filename is the class name. By default, the header is added if not specified
     * @deprecated since version 0.4.0
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static File objectsToExcel(List<?> objects, Class<?> clazz, Extension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException, SheetAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), extension, true);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. The default filename is the class name
     * @deprecated since version 0.4.0
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @param writeHeader If {@code true} it will write the header to the first line
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static File objectsToExcel(List<?> objects, Class<?> clazz, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException, SheetAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), clazz.getSimpleName(), extension, writeHeader);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder. By default, the header is added if not specified
     * @deprecated since version 0.4.0
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param filename The name of the output file without the extension
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String filename, Extension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException, SheetAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, extension, true);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}<p>
     * The default path is that of the temporary folder
     * @deprecated since version 0.4.0
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
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String filename, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException, SheetAlreadyExistsException {
        return objectsToExcel(objects, clazz, System.getProperty("java.io.tmpdir"), filename, extension, writeHeader);
    }

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}
     * @deprecated since version 0.4.0
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
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException, SheetAlreadyExistsException {
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
        ExcelWorkbook excelWorkbook = ExcelWorkbook.create(extension);
        Workbook workbook = excelWorkbook.getWorkbook();
        objectsToExistingExcel(workbook, objects, clazz, writeHeader);

        /* Write file */
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);

        /* Close file */
        excelWorkbook.close(fileOutputStream);

        return file;
    }

    /**
     * @param objectToExcels A list of {@code ObjectToExcels}. Each element represents a Sheet
     * @param extension The extension to assign to the Excel file
     * @param filename The name, including the path, to be associated with the File.<p> Note: Do not add the extension
     * @param writeHeader If {@code true} it will write the header to the first line
     * @return The Excel file generated by the conversion
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     * @since 0.4.0
     */
    public static File objectsToExcelFile(List<ObjectToExcel<?>> objectToExcels, Extension extension, String filename, Boolean writeHeader) throws ExtensionNotValidException, IOException, FileAlreadyExistsException, SheetAlreadyExistsException {
        File file = new File(filename + "." + extension.getExt());
        if (file.exists())
            throw new FileAlreadyExistsException("There is already a file with this pathname: " + file.getAbsolutePath());

        byte[] byteResult = objectsToExcelByte(objectToExcels, extension, writeHeader);
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        fileOutputStream.write(byteResult);
        fileOutputStream.close();

        return file;
    }

    /**
     * @param objectToExcels A list of {@code ObjectToExcels}. Each element represents a Sheet
     * @param extension The extension to assign to the Excel file
     * @param writeHeader If {@code true} it will write the header to the first line
     * @return The Excel OutputStream generated by the conversion in the form of a byte array
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     * @since 0.4.0
     */
    public static byte[] objectsToExcelByte(List<ObjectToExcel<?>> objectToExcels, Extension extension, Boolean writeHeader) throws ExtensionNotValidException, IOException, SheetAlreadyExistsException {
        ByteArrayOutputStream outputStream = objectsToExcelStream(objectToExcels, extension, writeHeader);
        return outputStream.toByteArray();
    }

    /**
     * @param objectToExcels A list of {@code ObjectToExcels}. Each element represents a Sheet
     * @param extension The extension to assign to the Excel file
     * @param writeHeader If {@code true} it will write the header to the first line
     * @return The Excel OutputStream generated by the conversion in the form of a ByteArrayOutputStream
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     * @since 0.4.0
     */
    public static ByteArrayOutputStream objectsToExcelStream(List<ObjectToExcel<?>> objectToExcels, Extension extension, Boolean writeHeader) throws ExtensionNotValidException, IOException, SheetAlreadyExistsException {
        /* Check extension*/
        if(!extension.isExcelExtension())
            throw new ExtensionNotValidException("Select an extension for an Excel file");

        /* Create workbook */
        ExcelWorkbook excelWorkbook = ExcelWorkbook.create(extension);

        /* Create a Sheet for each element */
        for(ObjectToExcel<?> objectToExcel : objectToExcels) {
            ExcelSheet excelSheet = excelWorkbook.createSheet(objectToExcel.getSheetName());
            Class<?> clazz = objectToExcel.getClazz();
            Field[] fields = clazz.getDeclaredFields();
            setFieldsAccessible(fields);
            AtomicInteger nRow = new AtomicInteger();

            /* Write header */
            if (writeHeader) {
                CellStyle headerCellStyle = createHeaderCellStyle(excelWorkbook, clazz);
                ExcelRow headerRow = excelSheet.createRow(nRow.getAndIncrement());
                for (int i = 0; i < fields.length; i++) {
                    ExcelCell excelCell = headerRow.createCell(i);
                    excelCell.getCell().setCellStyle(headerCellStyle);
                    ExcelField excelField = fields[i].getAnnotation(ExcelField.class);
                    excelCell.writeValue(excelField != null ? excelField.name() : fields[i].getName());
                }
            }

            /* Write body */
            objectToExcel.getStream().forEach(object -> {
                CellStyle bodyCellStyle = createBodyStyle(excelWorkbook, clazz);
                ExcelRow excelRow = excelSheet.createRow(nRow.getAndIncrement());
                for (int i = 0; i < fields.length; i++) {
                    ExcelCell excelCell = excelRow.createCell(i);
                    excelCell.getCell().setCellStyle(bodyCellStyle);
                    try {
                        excelCell.writeValue(fields[i].get(object));
                    } catch (IllegalAccessException e) {
                        throw new RuntimeException(e);
                    }
                }
            });

            /* Set auto-size columns */
            setAutoSizeColumn(excelSheet, fields, clazz);
        }

        /* Write and close */
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        excelWorkbook.writeAndClose(outputStream);
        return outputStream;
    }

    /**
     * This method allows you to convert objects into a Sheet of a File that already exists.<p>
     * By default, the header is added if not specified
     * @deprecated since version 0.4.0
     * @param file The {@code File} to update
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     * @since 0.2.1
     */
    @Deprecated
    public static void objectsToExistingExcel(File file, List<?> objects, Class<?> clazz) throws OpenWorkbookException, ExtensionNotValidException, IOException, IllegalAccessException, SheetAlreadyExistsException {
        objectsToExistingExcel(file, objects, clazz, true);
    }

    /**
     * This method allows you to convert objects into a Sheet of a File that already exists.
     * @deprecated since version 0.4.0
     * @param file The {@code File} to update
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param writeHeader If {@code true} it will write the header to the first line
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     * @since 0.2.1
     */
    @Deprecated
    public static void objectsToExistingExcel(File file, List<?> objects, Class<?> clazz, Boolean writeHeader) throws OpenWorkbookException, ExtensionNotValidException, IOException, IllegalAccessException, SheetAlreadyExistsException {
        /* Open workbook */
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(file);
        Workbook workbook = excelWorkbook.getWorkbook();
        objectsToExistingExcel(workbook, objects, clazz, writeHeader);

        /* Write file */
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);

        /* Close file */
        excelWorkbook.close(fileOutputStream);
    }

    /**
     * This method allows you to convert objects into a Sheet of a Workbook that already exists.<p>
     * Note: This method does not call the "write" method of the workbook.<p>
     * By default, the header is added if not specified
     * @deprecated since version 0.4.0
     * @param workbook The {@code Workbook} to update
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static void objectsToExistingExcel(Workbook workbook, List<?> objects, Class<?> clazz) throws IllegalAccessException, SheetAlreadyExistsException {
        objectsToExistingExcel(workbook, objects, clazz, true);
    }

    /**
     * This method allows you to convert objects into a Sheet of a Workbook that already exists.<p>
     * Note: This method does not call the "write" method of the workbook.
     * @deprecated since version 0.4.0
     * @param workbook The {@code Workbook} to update
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param writeHeader If {@code true} it will write the header to the first line
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     */
    @Deprecated
    public static void objectsToExistingExcel(Workbook workbook, List<?> objects, Class<?> clazz, Boolean writeHeader) throws IllegalAccessException, SheetAlreadyExistsException {
        /* Create sheet */
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(workbook);
        Sheet sheet = excelWorkbook.createSheet(clazz.getSimpleName()).getSheet();

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
     * @deprecated since version 0.4.0
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
    @Deprecated
    public static List<?> excelToObjects(File file, Class<?> clazz) throws ExtensionNotValidException, IOException, OpenWorkbookException, InstantiationException, IllegalAccessException, InvocationTargetException, NoSuchMethodException, SheetNotFoundException, HeaderNotPresentException {
        return excelToObjects(file, clazz, null);
    }

    /**
     * Convert an Excel file into a list of objects<p>
     * Note: The type of the elements of the return objects must coincide with the type of {@code clazz}
     * @deprecated since version 0.4.0
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
    @Deprecated
    public static List<?> excelToObjects(File file, Class<?> clazz, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, InvocationTargetException, IllegalAccessException, NoSuchMethodException, InstantiationException, SheetNotFoundException, HeaderNotPresentException {
        /* Open file excel */
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(file);
        Sheet sheet = (sheetName == null || sheetName.isEmpty())
                ? excelWorkbook.getSheet(0).getSheet()
                : excelWorkbook.getSheet(sheetName).getSheet();

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
        excelWorkbook.close();

        return resultList;
    }

    /**
     * @param bytes The byte array of the Excel file
     * @param excelToObjects A list of {@code ExcelToObject}. Each element represents a Sheet
     * @return A map where the key represents the name of the Sheet and the value the Stream of objects present in the sheet
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws ReadValueException If an error occurs while reading a cell
     * @throws HeaderNotPresentException If the first row is empty and does not contain the header
     * @throws InvocationTargetException If an error occurs while instantiating a new object or setting a field
     * @throws NoSuchMethodException If the setting method or empty constructor of the object is not found
     * @throws InstantiationException f an error occurs while instantiating a new object
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error occurs
     * @since 0.4.0
     */
    public static Map<String, Stream<?>> excelByteToObjects(byte[] bytes, List<ExcelToObject<?>> excelToObjects) throws OpenWorkbookException, SheetNotFoundException, ReadValueException, HeaderNotPresentException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException, IOException {
        ByteArrayInputStream byteStream = new ByteArrayInputStream(bytes);
        return excelStreamToObjects(byteStream, excelToObjects);
    }

    /**
     * @param file The Excel file
     * @param excelToObjects A list of {@code ExcelToObject}. Each element represents a Sheet
     * @return A map where the key represents the name of the Sheet and the value the Stream of objects present in the sheet
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws ReadValueException If an error occurs while reading a cell
     * @throws HeaderNotPresentException If the first row is empty and does not contain the header
     * @throws InvocationTargetException If an error occurs while instantiating a new object or setting a field
     * @throws NoSuchMethodException If the setting method or empty constructor of the object is not found
     * @throws InstantiationException f an error occurs while instantiating a new object
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @since 0.4.0
     */
    public static Map<String, Stream<?>> excelFileToObjects(File file, List<ExcelToObject<?>> excelToObjects) throws IOException, OpenWorkbookException, SheetNotFoundException, ReadValueException, HeaderNotPresentException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException, ExtensionNotValidException {
        if (!ExcelUtility.isValidExcelExtension(FilenameUtils.getExtension(file.getName()))) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }
        FileInputStream fileInputStream = new FileInputStream(file);
        Map<String, Stream<?>> map = excelStreamToObjects(fileInputStream, excelToObjects);
        fileInputStream.close();
        return map;
    }

    /**
     * @param inputStream The InputStream of the Excel file
     * @param excelToObjects A list of {@code ExcelToObject}. Each element represents a Sheet
     * @return A map where the key represents the name of the Sheet and the value the Stream of objects present in the sheet
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws HeaderNotPresentException If the first row is empty and does not contain the header
     * @throws NoSuchMethodException If the setting method or empty constructor of the object is not found
     * @throws InvocationTargetException If an error occurs while instantiating a new object or setting a field
     * @throws InstantiationException f an error occurs while instantiating a new object
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws ReadValueException If an error occurs while reading a cell
     * @throws IOException If an I/O error occurs
     * @since 0.4.0
     */
    public static Map<String, Stream<?>> excelStreamToObjects(InputStream inputStream, List<ExcelToObject<?>> excelToObjects) throws OpenWorkbookException, SheetNotFoundException, HeaderNotPresentException, NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException, ReadValueException, IOException {
        /* Open Workbook */
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(inputStream);
        Map<String, Stream<?>> map = new HashMap<>();

        /* Iterate all the sheets to convert */
        for (ExcelToObject<?> excelToObject : excelToObjects) {
            Class<?> clazz = excelToObject.getClazz();
            ExcelSheet excelSheet = excelWorkbook.getSheet(excelToObject.getSheetName());

            /* Retrieving header names */
            Field[] fields = clazz.getDeclaredFields();
            setFieldsAccessible(fields);
            Map<Integer, String> headerMap = getHeaderNames(excelSheet, fields);

            Stream.Builder<Object> streamBuilder = Stream.builder();

            /* Iterate all rows */
            for (ExcelRow excelRow : excelSheet.getRows()) {
                if (excelRow.getRow() == null || excelRow.getIndex() == 0) {
                    continue;
                }

                Object obj = clazz.getDeclaredConstructor().newInstance();
                /* Iterate all cells */
                for (ExcelCell excelCell : excelRow.getCells()) {
                    if (excelCell.getCell() == null) {
                        continue;
                    }

                    String headerName = headerMap.get(excelCell.getIndex());
                    if (headerName == null || headerMap.isEmpty()) {
                        continue;
                    }

                    /* Read the value in the cell */
                    Optional<Field> fieldOptional = Arrays.stream(fields).filter(f -> f.getName().equalsIgnoreCase(headerName)).findFirst();
                    if (fieldOptional.isEmpty()) {
                        throw new RuntimeException();
                    }
                    Field field = fieldOptional.get();

                    /* Set the value */
                    String methodName = "set" + Character.toUpperCase(headerName.charAt(0)) + headerName.substring(1);
                    Method setMethod = clazz.getMethod(methodName, field.getType());
                    setMethod.invoke(obj, excelCell.readValue(field.getType()));
                }
                /* Adds the object to the Stream after it has finished cycling through all cells */
                streamBuilder.add(obj);
            }
            /* inserts the Stream of all Sheet objects into the Map */
            map.put(excelSheet.getName(), streamBuilder.build());
        }

        excelWorkbook.close(inputStream);
        return map;
    }

    /**
     * Convert an Excel file into a CSV file<p>
     * The default path is that of the temporary folder. By default, the first sheet is chosen and the filename will be the same as the input file if not specified
     * @deprecated since version 0.4.0
     * @param fileInput The input Excel file that will be converted into a CSV file
     * @return A CSV file that contains the same lines as the Excel file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     */
    @Deprecated
    public static File excelToCsv(File fileInput) throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        return excelToCsv(fileInput, System.getProperty("java.io.tmpdir"), fileInput.getName().split("\\.")[0].trim(), null);
    }

    /**
     * Convert an Excel file into a CSV file<p>
     * The default path is that of the temporary folder. By default, the first sheet is chosen and the filename will be the same as the input file if not specified
     * @deprecated since version 0.4.0
     * @param fileInput The input Excel file that will be converted into a CSV file
     * @param sheetName The name of the sheet to open
     * @return A CSV file that contains the same lines as the Excel file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     */
    @Deprecated
    public static File excelToCsv(File fileInput, String sheetName) throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        return excelToCsv(fileInput, System.getProperty("java.io.tmpdir"), fileInput.getName().split("\\.")[0].trim(), sheetName);
    }

    /**
     * Convert an Excel file into a CSV file<p>
     * By default, the first sheet is chosen
     * @deprecated since version 0.4.0
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
    @Deprecated
    public static File excelToCsv(File fileInput, String path, String filename) throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        return excelToCsv(fileInput, path, filename, null);
    }

    /**
     * Convert an Excel file into a CSV file
     * @deprecated since version 0.4.0
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
    @Deprecated
    public static File excelToCsv(File fileInput, String path, String filename, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException, FileAlreadyExistsException {
        /* Check the Excel extension */
        if (!ExcelUtility.isValidExcelExtension(FilenameUtils.getExtension(fileInput.getName()))) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }

        /* Open file excel */
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(fileInput);
        Sheet sheet = (sheetName == null || sheetName.isEmpty())
                ? excelWorkbook.getSheet(0).getSheet()
                : excelWorkbook.getSheet(sheetName).getSheet();

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
        excelWorkbook.close(csvWriter);

        return csvFile;
    }

    /**
     * @param bytes The Excel file in the form of a byte array
     * @return A map where the key represents the Sheet name and the value is a CSV file for each Sheet
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws IOException If an I/O error has occurred
     * @since 0.4.0
     */
    public static Map<String, byte[]> excelToCsvByte(byte[] bytes) throws OpenWorkbookException, IOException {
        /* Open InputStream */
        InputStream inputStream = new ByteArrayInputStream(bytes);
        Map<String, byte[]> byteArrayMap = new HashMap<>();

        Map<String, ByteArrayOutputStream> outputStreamMap = excelToCsvStream(inputStream);

        /* iterate all the outputStream */
        for (Map.Entry<String, ByteArrayOutputStream> entry : outputStreamMap.entrySet()) {
            ByteArrayOutputStream baos = entry.getValue();
            byteArrayMap.put(entry.getKey(), baos.toByteArray());
        }

        return byteArrayMap;
    }

    /**
     * @param excelFile The Excel file
     * @param path The path, without file name, where the files will be saved
     * @return A map where the key represents the Sheet name and the value is a CSV file for each Sheet
     * @throws IOException If an I/O error has occurred
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws ExtensionNotValidException If an extension is not correct
     * @since 0.4.0
     */
    public static Map<String, File> excelToCsvFile(File excelFile, String path) throws IOException, OpenWorkbookException, ExtensionNotValidException {
        if (!ExcelUtility.isValidExcelExtension(FilenameUtils.getExtension(excelFile.getName()))) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }
        /* Open InputStream */
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        Map<String, File> fileMap = new HashMap<>();

        Map<String, ByteArrayOutputStream> outputStreamMap = excelToCsvStream(fileInputStream);

        /* iterate all the outputStream */
        for (Map.Entry<String, ByteArrayOutputStream> entry : outputStreamMap.entrySet()) {
            String pathname = path + "/" + entry.getKey() + "." + Extension.CSV.getExt();
            FileOutputStream fileOutputStream = new FileOutputStream(pathname);
            ByteArrayOutputStream baos = entry.getValue();
            fileOutputStream.write(baos.toByteArray());
            File file = new File(pathname);
            fileMap.put(entry.getKey(), file);
            fileOutputStream.close();
        }

        return fileMap;
    }

    /**
     * @param excelStream The Excel file in the form of an InputStream
     * @return A map where the key represents the Sheet name and the value is a CSV file for each Sheet
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws IOException If an I/O error has occurred
     * @since 0.4.0
     */
    public static Map<String, ByteArrayOutputStream> excelToCsvStream(InputStream excelStream) throws OpenWorkbookException, IOException {
        /* Open file excel */
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(excelStream);
        List<ExcelSheet> excelSheets = excelWorkbook.getSheets();

        Map<String, ByteArrayOutputStream> map = new HashMap<>();

        /* Iterate all the Sheets */
        for (ExcelSheet excelSheet : excelSheets) {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

            /* Open CSV Writer */
            Writer writer = new OutputStreamWriter(outputStream);
            CSVWriter csvWriter = new CSVWriter(writer);

            for (ExcelRow excelRow : excelSheet.getRows()) {
                List<String> data = new LinkedList<>();
                for (ExcelCell excelCell : excelRow.getCells()) {
                    data.add(excelCell.readValueAsString());
                }
                csvWriter.writeNext(data.toArray(data.toArray(new String[0])));
            }
            csvWriter.close();
            map.put(excelSheet.getName(), outputStream);
        }

        /* Close workbook */
        excelWorkbook.close(excelStream);

        return map;
    }

    /**
     * Convert a CSV file into an Excel file<p>
     * The default path is that of the temporary folder. By default, the filename will be the same as the input file if not specified and the extension is XLSX
     * @deprecated since version 0.4.0
     * @param fileInput The input CSV file that will be converted into an Excel file
     * @return An Excel file that contains the same lines as the CSV file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     * @throws IOException If an I/O error has occurred
     */
    @Deprecated
    public static File csvToExcel(File fileInput) throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException {
        return csvToExcel(fileInput, System.getProperty("java.io.tmpdir"), fileInput.getName().split("\\.")[0].trim(), Extension.XLSX);
    }

    /**
     * Convert a CSV file into an Excel file<p>
     * The default path is that of the temporary folder. By default, the extension is XLSX
     * @deprecated since version 0.4.0
     * @param fileInput The input CSV file that will be converted into an Excel file
     * @param filename The name of the output file without the extension
     * @return An Excel file that contains the same lines as the CSV file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     * @throws IOException If an I/O error has occurred
     */
    @Deprecated
    public static File csvToExcel(File fileInput, String filename) throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException {
        return csvToExcel(fileInput, System.getProperty("java.io.tmpdir"), filename, Extension.XLSX);
    }

    /**
     * Convert a CSV file into an Excel file<p>
     * By default, the extension is XLSX
     * @deprecated since version 0.4.0
     * @param fileInput The input CSV file that will be converted into an Excel file
     * @param path The destination path of the output file
     * @param filename The name of the output file without the extension
     * @return An Excel file that contains the same lines as the CSV file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     * @throws IOException If an I/O error has occurred
     */
    @Deprecated
    public static File csvToExcel(File fileInput, String path, String filename) throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException {
        return csvToExcel(fileInput, path, filename, Extension.XLSX);
    }

    /**
     * Convert a CSV file into an Excel file
     * @deprecated since version 0.4.0
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
    @Deprecated
    public static File csvToExcel(File fileInput, String path, String filename, Extension extension) throws IOException, ExtensionNotValidException, CsvValidationException, FileAlreadyExistsException {
        /* Check the Excel extension */
        if (!ExcelUtility.isValidExcelExtension(extension.getExt())) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }

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
        ExcelWorkbook excelWorkbook = ExcelWorkbook.create(extension);
        Workbook workbook = excelWorkbook.getWorkbook();
        csvToExistingExcel(workbook, csvReader);

        /* Write file */
        FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
        workbook.write(fileOutputStream);

        /* Close file */
        excelWorkbook.close(fileOutputStream, csvReader);

        return outputFile;
    }

    /**
     * @param bytes The CSV file in the form of a byte array
     * @param sheetName The name of the sheet to create
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @return The Excel file in the form of a byte array
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If an extension is not correct
     * @throws IOException If an I/O error has occurred
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     * @since 0.4.0
     */
    public static byte[] csvToExcelByte(byte[] bytes, String sheetName, Extension extension) throws CsvValidationException, ExtensionNotValidException, IOException, SheetAlreadyExistsException {
        InputStream inputStream = new ByteArrayInputStream(bytes);
        ByteArrayOutputStream baos = csvToExcelStream(inputStream, sheetName, extension);
        return baos.toByteArray();
    }

    /**
     * @param fileInput The CSV file
     * @param sheetName The name of the sheet to create
     * @param pathname The path with the file name. Note: Do not add the extension
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @return The Excel file
     * @throws IOException If an I/O error has occurred
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If an extension is not correct
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     * @since 0.4.0
     */
    public static File csvToExcelFile(File fileInput, String sheetName, String pathname, Extension extension) throws IOException, CsvValidationException, ExtensionNotValidException, SheetAlreadyExistsException {
        isValidCsvExtension(FilenameUtils.getExtension(fileInput.getName()));
        InputStream inputStream = new FileInputStream(fileInput);
        ByteArrayOutputStream baos = csvToExcelStream(inputStream, sheetName, extension);
        pathname = pathname + "." + extension.getExt();
        FileOutputStream fileOutputStream = new FileOutputStream(pathname);
        fileOutputStream.write(baos.toByteArray());
        File file = new File(pathname);
        fileOutputStream.close();

        return file;
    }

    /**
     * @param inputStream The CSV file in the form of a InputStream
     * @param sheetName The name of the sheet to create
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @return The Excel file in the form of a ByteArrayOutputStream
     * @throws ExtensionNotValidException If an extension is not correct
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws IOException If an I/O error has occurred
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     * @since 0.4.0
     */
    public static ByteArrayOutputStream csvToExcelStream(InputStream inputStream, String sheetName, Extension extension) throws ExtensionNotValidException, CsvValidationException, IOException, SheetAlreadyExistsException {
        /* Check the extension */
        if (!ExcelUtility.isValidExcelExtension(extension.getExt())) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }

        /* Open CSV Reader */
        Reader reader = new InputStreamReader(inputStream);
        CSVReader csvReader = new CSVReader(reader);

        ExcelWorkbook excelWorkbook = new ExcelWorkbook(extension);
        ExcelSheet excelSheet = excelWorkbook.createSheet(sheetName);

        /* Read CSV file */
        String[] values;
        int cRow = 0;
        while ((values = csvReader.readNext()) != null) {
            ExcelRow excelRow = excelSheet.createRow(cRow);
            for (int j = 0; j < values.length; j++) {
                ExcelCell excelCell = excelRow.createCell(j);
                excelCell.writeValue(values[j]);
                excelSheet.getSheet().autoSizeColumn(j);
            }
            cRow++;
        }

        /* Write and close the Workbook */
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        inputStream.close();
        excelWorkbook.writeAndClose(outputStream);
        csvReader.close();

        return outputStream;
    }

    /**
     * Convert the CSV file into a new sheet of an existing File.
     * @deprecated since version 0.4.0
     * @param fileOutput The {@code File} to update
     * @param fileInput The input CSV file that will be converted into an Excel file
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     * @throws IOException If an I/O error has occurred
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @since 0.2.1
     */
    @Deprecated
    public static void csvToExistingExcel(File fileOutput, File fileInput) throws OpenWorkbookException, ExtensionNotValidException, IOException, CsvValidationException {
        /* Check the Excel extension */
        if (!ExcelUtility.isValidExcelExtension(FilenameUtils.getExtension(fileOutput.getName()))) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }

        /* Check the CSV extension */
        isValidCsvExtension(FilenameUtils.getExtension(fileInput.getName()));

        /* Open workbook */
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(fileOutput);
        Workbook workbook = excelWorkbook.getWorkbook();
        csvToExistingExcel(workbook, fileInput);

        /* Write file */
        FileOutputStream fileOutputStream = new FileOutputStream(fileOutput);
        workbook.write(fileOutputStream);

        /* Close file */
        excelWorkbook.close(fileOutputStream);
    }

    /**
     * Writes the data present in the CSVReader to a new sheet of an existing File.
     * @deprecated since version 0.4.0
     * @param fileOutput The {@code File} to update
     * @param csvReader The {@code CSVReader} of the CSV input file
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     * @throws IOException If an I/O error has occurred
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @since 0.2.1
     */
    @Deprecated
    public static void csvToExistingExcel(File fileOutput, CSVReader csvReader) throws OpenWorkbookException, ExtensionNotValidException, IOException, CsvValidationException {
        /* Check the Excel extension */
        if (!ExcelUtility.isValidExcelExtension(FilenameUtils.getExtension(fileOutput.getName()))) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }

        /* Open workbook */
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(fileOutput);
        Workbook workbook = excelWorkbook.getWorkbook();
        csvToExistingExcel(workbook, csvReader);

        /* Write file */
        FileOutputStream fileOutputStream = new FileOutputStream(fileOutput);
        workbook.write(fileOutputStream);

        /* Close file */
        excelWorkbook.close(fileOutputStream, csvReader);
    }

    /**
     * Convert the CSV file into a new sheet of an existing Workbook.<p>
     * Note: This method does not call the "write" method of the workbook.
     * @deprecated since version 0.4.0
     * @param workbook The {@code Workbook} to update
     * @param fileInput The input CSV file that will be converted into an Excel file
     * @throws IOException If an I/O error has occurred
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     */
    @Deprecated
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
     * @deprecated since version 0.4.0
     * @param workbook The {@code Workbook} to update
     * @param csvReader The {@code CSVReader} of the CSV input file
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws IOException If an I/O error has occurred
     */
    @Deprecated
    public static void csvToExistingExcel(Workbook workbook, CSVReader csvReader) throws CsvValidationException, IOException {
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(workbook);
        ExcelSheet excelSheet = excelWorkbook.createSheet();
        Sheet sheet = excelSheet.getSheet();

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

    /**
     * @param bytes The Excel file in the form of a byte array
     * @return The Json file in the form of a byte array
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws IOException If an I/O error has occurred
     * @since 0.4.0
     */
    public static byte[] excelToJsonByte(byte[] bytes) throws OpenWorkbookException, IOException {
        InputStream inputStream = new ByteArrayInputStream(bytes);
        ByteArrayOutputStream baos = excelToJsonStream(inputStream);
        return baos.toByteArray();
    }

    /**
     * @param excelFile The Excel file
     * @param pathname The path with the file name. Note: Do not add the extension
     * @return The Json file
     * @throws IOException If an I/O error has occurred
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws ExtensionNotValidException If the extension is incorrect
     * @since 0.4.0
     */
    public static File excelToJsonFile(File excelFile, String pathname) throws IOException, OpenWorkbookException, ExtensionNotValidException {
        if (!ExcelUtility.isValidExcelExtension(FilenameUtils.getExtension(excelFile.getName()))) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }
        FileInputStream fileInputStream = new FileInputStream(excelFile);
        ByteArrayOutputStream baos = excelToJsonStream(fileInputStream);
        pathname = pathname + "." + Extension.JSON.getExt();
        FileOutputStream fileOutputStream = new FileOutputStream(pathname);
        fileOutputStream.write(baos.toByteArray());
        File jsonFile = new File(pathname);
        fileOutputStream.close();

        return jsonFile;
    }

    /**
     * @param excelStream The Excel file in the form of an InputStream
     * @return The Json file in the form of a ByteArrayOutputStream
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws IOException If an I/O error has occurred
     * @since 0.4.0
     */
    public static ByteArrayOutputStream excelToJsonStream(InputStream excelStream) throws OpenWorkbookException, IOException {
        /* Open Workbook */
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(excelStream);
        List<ExcelSheet> excelSheets = excelWorkbook.getSheets();

        /* Create Json Object */
        ObjectMapper objectMapper = new ObjectMapper();
        ObjectNode rootNode = objectMapper.createObjectNode();

        /* Iterate all the Sheets */
        for (ExcelSheet excelSheet : excelSheets) {
            /* Create a node for each Sheet */
            ObjectNode sheetNode = objectMapper.createObjectNode();
            List<ExcelRow> excelRows = excelSheet.getRows();

            /* Iterate all the Rows */
            for (ExcelRow excelRow : excelRows) {
                /* Create a node for each Row */
                ObjectNode rowNode = objectMapper.createObjectNode();
                List<ExcelCell> excelCells = excelRow.getCells();

                /* Iterate all the Cells */
                for (ExcelCell excelCell : excelCells) {
                    rowNode.put("col_" + (excelCell.getIndex() + 1), excelCell.readValueAsString());
                }
                sheetNode.putPOJO("row_" + (excelRow.getIndex() + 1), rowNode);
            }

            rootNode.putPOJO(excelSheet.getName(), sheetNode);
        }

        /* Close Workbook */
        excelWorkbook.close(excelStream);

        /* Create OutputStream */
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        objectMapper.writeValue(byteArrayOutputStream, rootNode);

        return byteArrayOutputStream;
    }

    /**
     * @param bytes The Json file in the form of a byte array. Note: Must be formatted as an array of objects
     * @param jsonToExcel It is used to convert Json elements into objects
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @param writeHeader If {@code true} it will write the header to the first line
     * @param <T> The class parameter of the object
     * @return The Excel file in the form of a byte array
     * @throws ExtensionNotValidException If the extension is incorrect
     * @throws IOException If an I/O error has occurred
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     * @since 0.4.0
     */
    public static <T> byte[] jsonToExcelByte(byte[] bytes, JsonToExcel<T> jsonToExcel, Extension extension, Boolean writeHeader) throws ExtensionNotValidException, IOException, SheetAlreadyExistsException {
        InputStream inputStream = new ByteArrayInputStream(bytes);
        return jsonToExcelStream(inputStream, jsonToExcel, extension, writeHeader).toByteArray();
    }

    /**
     * @param jsonFile The Json file. Note: Must be formatted as an array of objects
     * @param jsonToExcel It is used to convert Json elements into objects
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @param filename File name, including path, to specify where to save. Note: Do not add the extension
     * @param writeHeader If {@code true} it will write the header to the first line
     * @param <T> The class parameter of the object
     * @return The Excel file
     * @throws IOException If an I/O error has occurred
     * @throws ExtensionNotValidException If the extension is incorrect
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     * @since 0.4.0
     */
    public static <T> File jsonToExcelFile(File jsonFile, JsonToExcel<T> jsonToExcel, Extension extension, String filename, Boolean writeHeader) throws IOException, ExtensionNotValidException, FileAlreadyExistsException, SheetAlreadyExistsException {
        /* Check extension */
        isValidJsonExtension(FilenameUtils.getExtension(jsonFile.getName()));
        FileInputStream fileInputStream = new FileInputStream(jsonFile);

        /* Convert the Json to a Stream */
        List<ObjectToExcel<?>> objectToExcels = jsonInputStreamToObjectToExcel(fileInputStream, jsonToExcel);
        fileInputStream.close();

        /* Call the method that converts objects to in Excel */
        return objectsToExcelFile(objectToExcels, extension, filename, writeHeader);
    }

    /**
     * @param jsonStream The Json file in the form of an InputStream. Note: Must be formatted as an array of objects
     * @param jsonToExcel It is used to convert Json elements into objects
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @param writeHeader If {@code true} it will write the header to the first line
     * @param <T> The class parameter of the object
     * @return The Excel file in the form of a ByteArrayOutputStream
     * @throws ExtensionNotValidException If the extension is incorrect
     * @throws IOException If an I/O error has occurred
     * @throws SheetAlreadyExistsException If you try to insert a Sheet that already exists
     * @since 0.4.0
     */
    public static <T> ByteArrayOutputStream jsonToExcelStream(InputStream jsonStream, JsonToExcel<T> jsonToExcel, Extension extension, Boolean writeHeader) throws ExtensionNotValidException, IOException, SheetAlreadyExistsException {
        /* Check extension */
        if(!extension.isExcelExtension())
            throw new ExtensionNotValidException("Select an extension for an Excel file");

        /* Convert the Json to a Stream */
        List<ObjectToExcel<?>> objectToExcels = jsonInputStreamToObjectToExcel(jsonStream, jsonToExcel);

        /* Call the method that converts objects to in Excel */
        return objectsToExcelStream(objectToExcels, extension, writeHeader);
    }

    private static <T> List<ObjectToExcel<?>> jsonInputStreamToObjectToExcel(InputStream jsonStream, JsonToExcel<T> jsonToExcel) throws IOException {
        Stream<T> jsonObjectsStream = jsonToStream(jsonStream, jsonToExcel.getClazz());
        ObjectToExcel<T> objectToExcel = new ObjectToExcel<>(jsonToExcel.getSheetName(), jsonToExcel.getClazz(), jsonObjectsStream);
        List<ObjectToExcel<?>> objectToExcels = new ArrayList<>();
        objectToExcels.add(objectToExcel);

        return objectToExcels;
    }

    private static <T> Stream<T> jsonToStream(InputStream inputStream, Class<T> clazz) throws IOException {
        ObjectMapper objectMapper = new ObjectMapper();
        JsonParser jsonParser = objectMapper.getFactory().createParser(inputStream);
        TypeReference<List<T>> typeRef = new TypeReference<>() {};
        List<T> hashList = objectMapper.readValue(jsonParser, typeRef);
        List<T> list = new ArrayList<>();
        for (T val : hashList) {
            T obj = objectMapper.convertValue(val, clazz);
            list.add(obj);
        }
        return list.stream();
    }

    private static void isValidCsvExtension(String extension) throws ExtensionNotValidException {
        if (!extension.equalsIgnoreCase(Extension.CSV.getExt()))
            throw new ExtensionNotValidException("Pass a file with the CSV extension");
    }

    private static void isValidJsonExtension(String extension) throws ExtensionNotValidException {
        if (!extension.equalsIgnoreCase(Extension.JSON.getExt()))
            throw new ExtensionNotValidException("Pass a file with the JSON extension");
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

    private static Map<Integer, String> getHeaderNames(ExcelSheet excelSheet, Field[] fields) throws HeaderNotPresentException {
        Map<String, String> fieldNames = new HashMap<>();
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            fieldNames.put(excelField == null ? field.getName() : excelField.name(), field.getName());
        }

        Row headerRow = excelSheet.getSheet().getRow(0);
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
                    } else {
                        logger.error("{} type is not supported. It was not possible to write '{}'", field.getType(), headerName);
                    }

                }
                case BOOLEAN -> PropertyUtils.setSimpleProperty(obj, headerName, cell.getBooleanCellValue());
                case STRING -> PropertyUtils.setSimpleProperty(obj, headerName, cell.getStringCellValue());
                default -> logger.error("Cell type not supported. It was not possible to write '{}'", headerName);
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

    private static CellStyle createHeaderCellStyle(ExcelWorkbook excelWorkbook, Class<?> clazz) {
        CellStyle cellStyle = excelWorkbook.getWorkbook().createCellStyle();
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

    private static CellStyle createBodyStyle(ExcelWorkbook excelWorkbook, Class<?> clazz) {
        CellStyle cellStyle = excelWorkbook.getWorkbook().createCellStyle();
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

    private static void setAutoSizeColumn(ExcelSheet excelSheet, Field[] fields, Class<?> clazz) {
        ExcelHeaderStyle excelHeaderStyle = clazz.getAnnotation(ExcelHeaderStyle.class);
        if (excelHeaderStyle != null && excelHeaderStyle.autoSize()) {
            for (int i = 0; i < fields.length; i++) {
                excelSheet.getSheet().autoSizeColumn(i);
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
