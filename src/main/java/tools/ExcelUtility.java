package tools;

import enums.Extension;
import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import tools.implementations.ExcelSheetUtilsImpl;
import tools.implementations.ExcelWorkbookUtilsImpl;
import tools.interfaces.ExcelSheetUtils;
import tools.interfaces.ExcelWorkbookUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

/**
 * {@code ExcelUtility} is the static class with the implementations of some utilities on Excel files
 * @author Mirko Benincasa
 * @since 0.2.0
 */
public class ExcelUtility {

    /**
     * Counts all rows in all sheets<p>
     * If not specified, empty lines will also be included
     * @param file an Excel file
     * @return A list with the number of rows present for each sheet
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    public static List<Integer> countAllRowsOfAllSheets(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        return countAllRowsOfAllSheets(file, true);
    }

    /**
     * * Counts all rows in all sheets
     * @param file an Excel file
     * @param alsoEmptyRows if {@code true} then it will also count rows with all empty cells
     * @return A list with the number of rows present for each sheet
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    public static List<Integer> countAllRowsOfAllSheets(File file, Boolean alsoEmptyRows) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Check extension */
        String extension = checkExcelExtension(file.getName());

        /* Open file excel */
        FileInputStream fileInputStream = new FileInputStream(file);
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        Workbook workbook = excelWorkbookUtils.open(fileInputStream, extension);

        List<Integer> values = new LinkedList<>();
        for (Sheet sheet : workbook) {
            if (alsoEmptyRows) {
                values.add(sheet.getLastRowNum() + 1);
                continue;
            }

            values.add(countOnlyRowsNotEmpty(sheet));
        }

        /* Close file */
        excelWorkbookUtils.close(workbook, fileInputStream);

        return values;
    }

    /**
     * Counts all rows in a sheet<p>
     * If not specified, empty lines will also be included
     * @param file An Excel file
     * @param sheetName The name of the sheet to open
     * @return A number that corresponds to all rows in the sheet
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     */
    public static Integer countAllRows(File file, String sheetName) throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        return countAllRows(file, sheetName, true);
    }

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
    public static Integer countAllRows(File file, String sheetName, Boolean alsoEmptyRows) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        /* Check extension */
        String extension = checkExcelExtension(file.getName());

        /* Open file excel */
        FileInputStream fileInputStream = new FileInputStream(file);
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        Workbook workbook = excelWorkbookUtils.open(fileInputStream, extension);
        ExcelSheetUtils excelSheetUtils = new ExcelSheetUtilsImpl();
        Sheet sheet = (sheetName == null || sheetName.isEmpty())
                ? excelSheetUtils.open(workbook)
                : excelSheetUtils.open(workbook, sheetName);

        /* Count all rows */
        int numRows = alsoEmptyRows
                ? sheet.getLastRowNum() + 1
                : countOnlyRowsNotEmpty(sheet);

        /* Close file */
        excelWorkbookUtils.close(workbook, fileInputStream);

        return numRows;
    }

    /**
     * Check if the extension is that of an Excel file
     * @param filename The name of the file with extension
     * @return The name of the extension
     * @throws ExtensionNotValidException If the filename extension does not belong to an Excel file
     */
    public static String checkExcelExtension(String filename) throws ExtensionNotValidException {
        String extension = FilenameUtils.getExtension(filename);
        if (!isValidExcelExtension(extension)) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }
        return extension;
    }

    /**
     * Check if the extension is that of an Excel file
     * @param extension The file's extension
     * @return {@code true} if it is the extension of an Excel file
     */
    public static Boolean isValidExcelExtension(String extension) {
        return extension.equalsIgnoreCase(Extension.XLS.getExt()) || extension.equalsIgnoreCase(Extension.XLSX.getExt());
    }

    private static int countOnlyRowsNotEmpty(Sheet sheet) {
        int numRows = sheet.getLastRowNum() + 1;
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            boolean isEmptyRow = true;

            if (row == null) {
                numRows--;
                continue;
            }

            for (int j = 0; j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                if (cell != null) {
                    Object val;
                    switch (cell.getCellType()) {
                        case NUMERIC -> val = cell.getNumericCellValue();
                        case BOOLEAN -> val = cell.getBooleanCellValue();
                        default -> val = cell.getStringCellValue();
                    }
                    if (val != null) {
                        isEmptyRow = false;
                        break;
                    }
                }
            }

            if (isEmptyRow) {
                numRows--;
            }
        }

        return numRows;
    }
}