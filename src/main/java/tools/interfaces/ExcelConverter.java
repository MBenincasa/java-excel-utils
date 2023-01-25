package tools.interfaces;

import com.opencsv.exceptions.CsvValidationException;
import enums.Extension;
import exceptions.*;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.List;

/**
 * The {@code ExcelConverter} interface groups methods that convert objects or files to Excel files and vice versa
 * @deprecated since version 0.2.0. View here {@link tools.Converter}
 * @see tools.Converter
 * @author Mirko Benincasa
 * @since 0.1.0
 */
@Deprecated
public interface ExcelConverter {

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    File objectsToExcel(List<?> objects, Class<?> clazz) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException;

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param filename The name of the output file without the extension
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    File objectsToExcel(List<?> objects, Class<?> clazz, String filename) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException;

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}
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
    File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException;

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}
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
    File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException;

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param writeHeader If {@code true} it will write the header to the first line
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    File objectsToExcel(List<?> objects, Class<?> clazz, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException;

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}
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
    File objectsToExcel(List<?> objects, Class<?> clazz, String filename, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException;

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}
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
    File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, Extension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException;

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}
     * @param objects The list of objects that will be converted into an Excel file
     * @param clazz The class of the list elements
     * @param extension The extension of the output file. Select an extension with {@code type} EXCEL
     * @return An Excel file with as many rows as there are elements in the list.
     * @throws IllegalAccessException If a field or fields of the {@code clazz} could not be accessed
     * @throws IOException If an I/O error has occurred
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     */
    File objectsToExcel(List<?> objects, Class<?> clazz, Extension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException;

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}
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
    File objectsToExcel(List<?> objects, Class<?> clazz, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException;

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}
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
    File objectsToExcel(List<?> objects, Class<?> clazz, String filename, Extension extension) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException;

    /**
     * Convert a list of objects into an Excel file<p>
     * Note: The type of the elements of the {@code objects} list must coincide with the type of {@code clazz}
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
    File objectsToExcel(List<?> objects, Class<?> clazz, String filename, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException;

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
    File objectsToExcel(List<?> objects, Class<?> clazz, String path, String filename, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException, FileAlreadyExistsException, ExtensionNotValidException;

    /**
     * Convert an Excel file into a list of objects<p>
     * Note: The type of the elements of the return objects must coincide with the type of {@code clazz}
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
    List<?> excelToObjects(File file, Class<?> clazz) throws ExtensionNotValidException, IOException, OpenWorkbookException, InstantiationException, IllegalAccessException, InvocationTargetException, NoSuchMethodException, SheetNotFoundException, HeaderNotPresentException;

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
    List<?> excelToObjects(File file, Class<?> clazz, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, InvocationTargetException, IllegalAccessException, NoSuchMethodException, InstantiationException, SheetNotFoundException, HeaderNotPresentException;

    /**
     * Convert an Excel file into a CSV file
     * @param fileInput The input Excel file that will be converted into a CSV file
     * @return A CSV file that contains the same lines as the Excel file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     */
    File excelToCsv(File fileInput) throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException;

    /**
     * Convert an Excel file into a CSV file
     * @param fileInput The input Excel file that will be converted into a CSV file
     * @param sheetName The name of the sheet to open
     * @return A CSV file that contains the same lines as the Excel file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     */
    File excelToCsv(File fileInput, String sheetName) throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException;

    /**
     * Convert an Excel file into a CSV file
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
    File excelToCsv(File fileInput, String path, String filename) throws FileAlreadyExistsException, OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException;

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
    File excelToCsv(File fileInput, String path, String filename, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException, FileAlreadyExistsException;

    /**
     * Convert a CSV file into an Excel file
     * @param fileInput The input CSV file that will be converted into an Excel file
     * @return An Excel file that contains the same lines as the CSV file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     * @throws IOException If an I/O error has occurred
     */
    File csvToExcel(File fileInput) throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException;

    /**
     * Convert a CSV file into an Excel file
     * @param fileInput The input CSV file that will be converted into an Excel file
     * @param filename The name of the output file without the extension
     * @return An Excel file that contains the same lines as the CSV file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     * @throws IOException If an I/O error has occurred
     */
    File csvToExcel(File fileInput, String filename) throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException;

    /**
     * Convert a CSV file into an Excel file
     * @param fileInput The input CSV file that will be converted into an Excel file
     * @param path The destination path of the output file
     * @param filename The name of the output file without the extension
     * @return An Excel file that contains the same lines as the CSV file
     * @throws FileAlreadyExistsException If the destination file already exists
     * @throws CsvValidationException If the CSV file has invalid formatting
     * @throws ExtensionNotValidException If the input file extension does not belong to a CSV file
     * @throws IOException If an I/O error has occurred
     */
    File csvToExcel(File fileInput, String path, String filename) throws FileAlreadyExistsException, CsvValidationException, ExtensionNotValidException, IOException;

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
    File csvToExcel(File fileInput, String path, String filename, Extension extension) throws IOException, ExtensionNotValidException, CsvValidationException, FileAlreadyExistsException;
}
