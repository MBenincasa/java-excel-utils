package tools.interfaces;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * The {@code ExcelUtils} interface groups utility methods
 * @author Mirko Benincasa
 * @since 0.1.0
 */
public interface ExcelUtils {

    /**
     * Counts all rows in all sheets
     * @param file An Excel file
     * @return A list with the number of rows present for each sheet
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    List<Integer> countAllRowsOfAllSheets(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    /**
     * Counts all rows in all sheets
     * @param file An Excel file
     * @param alsoEmptyRows if {@code true} then it will also count rows with all empty cells
     * @return A list with the number of rows present for each sheet
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    List<Integer> countAllRowsOfAllSheets(File file, Boolean alsoEmptyRows) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    /**
     * Counts all rows in a sheet
     * @param file An Excel file
     * @param sheetName The name of the sheet to open
     * @return A number that corresponds to all rows in the sheet
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     */
    Integer countAllRows(File file, String sheetName) throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException;

    /**
     * Counts all rows in a sheet
     * @param file An Excel file
     * @param sheetName The name of the sheet to open
     * @param alsoEmptyRows if {@code true} then it will also count rows with all empty cells
     * @return A number that corresponds to all rows in the sheet
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     */
    Integer countAllRows(File file, String sheetName, Boolean alsoEmptyRows) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    /**
     * Check if the extension is that of an Excel file
     * @param extension The file's extension
     * @return {@code true} if it is the extension of an Excel file
     */
    Boolean isValidExcelExtension(String extension);

    /**
     * Check if the extension is that of an Excel file
     * @param filename The name of the file with extension
     * @return The name of the extension
     * @throws ExtensionNotValidException If the filename extension does not belong to an Excel file
     */
    String checkExcelExtension(String filename) throws ExtensionNotValidException;
}
