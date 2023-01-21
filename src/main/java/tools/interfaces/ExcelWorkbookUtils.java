package tools.interfaces;

import com.opencsv.CSVWriter;
import enums.Extension;
import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

/**
 * The {@code ExcelWorkbookUtils} interface groups methods that work with workbooks
 * @author Mirko Benincasa
 * @since 0.1.0
 */
public interface ExcelWorkbookUtils {

    /**
     * Opens the workbook
     * @param file An Excel file
     * @return The workbook that was opened
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    Workbook open(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    /**
     * Opens the workbook
     * @param fileInputStream The {@code FileInputStream} of the Excel file
     * @param extension The file's extension
     * @return The workbook that was opened
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    Workbook open(FileInputStream fileInputStream, String extension) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    /**
     * Create a new workbook
     * @return A workbook
     */
    Workbook create();

    /**
     * Create a new workbook
     * @param extension The extension of the file. Provide the extension of an Excel file
     * @return A workbook
     * @throws ExtensionNotValidException If the extension does not belong to an Excel file
     */
    Workbook create(String extension) throws ExtensionNotValidException;

    /**
     * Create a new workbook
     * @param extension The extension of the file. Select an extension with {@code type} EXCEL
     * @return A workbook
     */
    Workbook create(Extension extension);

    /**
     * Close a workbook
     * @param workbook The {@code Workbook} to close
     * @throws IOException If an I/O error has occurred
     */
    void close(Workbook workbook) throws IOException;

    /**
     * Close a workbook
     * @param workbook The {@code Workbook} to close
     * @param fileInputStream The {@code FileInputStream} to close
     * @throws IOException If an I/O error has occurred
     */
    void close(Workbook workbook, FileInputStream fileInputStream) throws IOException;

    /**
     * Close a workbook
     * @param workbook The {@code Workbook} to close
     * @param fileOutputStream The {@code FileOutputStream} to close
     * @throws IOException If an I/O error has occurred
     */
    void close(Workbook workbook, FileOutputStream fileOutputStream) throws IOException;

    /**
     * Close a workbook
     * @param workbook The {@code Workbook} to close
     * @param fileOutputStream The {@code FileOutputStream} to close
     * @param fileInputStream The {@code FileInputStream} to close
     * @throws IOException If an I/O error has occurred
     */
    void close(Workbook workbook, FileOutputStream fileOutputStream, FileInputStream fileInputStream) throws IOException;

    /**
     * Close a workbook
     * @param workbook The {@code Workbook} to close
     * @param fileInputStream The {@code FileInputStream} to close
     * @param writer The {@code CSVWriter} to close
     * @throws IOException If an I/O error has occurred
     */
    void close(Workbook workbook, FileInputStream fileInputStream, CSVWriter writer) throws IOException;
}
