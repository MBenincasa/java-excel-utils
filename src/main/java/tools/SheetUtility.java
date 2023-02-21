package tools;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;
import model.ExcelWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

/**
 * {@code SheetUtility} is the static class with implementations of some methods for working with Sheets
 * @author Mirko Benincasa
 * @since 0.2.0
 */
public class SheetUtility {

    /**
     * Count how many sheets there are
     * @param file Excel file
     * @return The number of sheets present
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    public static Integer length(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Open file excel */
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(file);
        Workbook workbook = excelWorkbook.getWorkbook();

        Integer totalSheets = workbook.getNumberOfSheets();

        /* Close file */
        excelWorkbook.close();

        return totalSheets;
    }

    /**
     * Returns the name of all sheets in the workbook
     * @param file Excel file
     * @return A list with the name of all sheets
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    public static List<String> getNames(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Open file excel */
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(file);
        Workbook workbook = excelWorkbook.getWorkbook();

        /* Iterate all the sheets */
        Iterator<Sheet> sheetIterator = workbook.iterator();
        List<String> sheetNames = new LinkedList<>();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            sheetNames.add(sheet.getSheetName());
        }

        /* Close file */
        excelWorkbook.close();

        return sheetNames;
    }

    /**
     * Search where the sheet is located by name
     * @param file Excel file
     * @param sheetName The name of the sheet
     * @return The position of the sheet
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    public static Integer getIndex(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        /* Open file excel */
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(file);
        Workbook workbook = excelWorkbook.getWorkbook();

        int sheetIndex = workbook.getSheetIndex(sheetName);

        /* Close file */
        excelWorkbook.close();

        if (sheetIndex < 0) {
            throw new SheetNotFoundException("No sheet was found");
        }
        return sheetIndex;
    }

    /**
     * Search for the sheet name by location
     * @param file Excel file
     * @param position The index in the workbook
     * @return The name of the sheet
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    public static String getName(File file, Integer position) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        /* Open file excel */
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(file);
        Workbook workbook = excelWorkbook.getWorkbook();

        String sheetName;
        try {
            sheetName = workbook.getSheetName(position);
        } catch (IllegalArgumentException e) {
            throw new SheetNotFoundException("Sheet index is out of range");
        }

        /* Close file */
        excelWorkbook.close();

        return sheetName;
    }

    /**
     * Create a sheet in a workbook
     * @param file Excel file
     * @return The new sheet that was created
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    public static Sheet create(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        return create(file, null);
    }

    /**
     * Create a sheet in a workbook
     * @param file Excel file
     * @param sheetName The name of the sheet to add
     * @return The new sheet that was created
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    public static Sheet create(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Check extension */
        String extension = ExcelUtility.checkExcelExtension(file.getName());

        /* Open file excel */
        FileInputStream fileInputStream = new FileInputStream(file);
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(fileInputStream, extension);
        Workbook workbook = excelWorkbook.getWorkbook();

        /* Create sheet */
        return sheetName == null ? workbook.createSheet() : workbook.createSheet(sheetName);
    }

    /**
     * Create a sheet in a workbook
     * @param workbook The {@code Workbook} where to add the new sheet
     * @return The new sheet that was created
     */
    public static Sheet create(Workbook workbook) {
        return create(workbook, null);
    }

    /**
     * Create a sheet in a workbook
     * @param workbook The {@code Workbook} where to add the new sheet
     * @param sheetName The name of the sheet to add
     * @return The new sheet that was created
     */
    public static Sheet create(Workbook workbook, String sheetName) {
        return sheetName == null ? workbook.createSheet() : workbook.createSheet(sheetName);
    }

    /**
     * Gets the sheet of the Excel file<p>
     * If not specified, the first sheet will be opened
     * @param file Excel file
     * @return The sheet in the workbook
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    public static Sheet get(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        return get(file, 0);
    }

    /**
     * Gets the sheet of the Excel file
     * @param file Excel file
     * @param sheetName The sheet name in the workbook
     * @return The sheet in the workbook
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    public static Sheet get(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        /* Open file excel */
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(file);
        Workbook workbook = excelWorkbook.getWorkbook();

        /* Open sheet */
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null)
            throw new SheetNotFoundException();

        /* Close workbook */
        excelWorkbook.close();

        return sheet;
    }

    /**
     * Gets the sheet of the Excel file
     * @param file Excel file
     * @param position The index in the workbook
     * @return The sheet in the workbook
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    public static Sheet get(File file, Integer position) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        /* Open file excel */
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(file);
        Workbook workbook = excelWorkbook.getWorkbook();

        /* Open sheet */
        Sheet sheet = workbook.getSheetAt(position);
        if (sheet == null)
            throw new SheetNotFoundException();

        /* Close workbook */
        excelWorkbook.close();

        return sheet;
    }

    /**
     * Gets the sheet in the workbook.<p>
     * If not specified, the first sheet will be opened
     * @param workbook The {@code Workbook} where there is the sheet
     * @return The sheet in the workbook in first position
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    public static Sheet get(Workbook workbook) throws SheetNotFoundException {
        return get(workbook, 0);
    }

    /**
     * Gets the sheet in the workbook.
     * @param workbook The {@code Workbook} where there is the sheet
     * @param sheetName The sheet name in the workbook
     * @return The sheet in the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    public static Sheet get(Workbook workbook, String sheetName) throws SheetNotFoundException {
        /* Open sheet */
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null)
            throw new SheetNotFoundException();
        return sheet;
    }

    /**
     * Gets the sheet in the workbook.
     * @param workbook The {@code Workbook} where there is the sheet
     * @param position The index in the workbook
     * @return The sheet in the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    public static Sheet get(Workbook workbook, Integer position) throws SheetNotFoundException {
        /* Open sheet */
        Sheet sheet = workbook.getSheetAt(position);
        if (sheet == null)
            throw new SheetNotFoundException();
        return sheet;
    }

    /**
     * Gets the sheet in the workbook. If it doesn't find it, it creates it.
     * @param workbook The {@code Workbook} where there is the sheet
     * @param sheetName The sheet name in the workbook
     * @return The sheet in the workbook or a new one
     */
    public static Sheet getOrCreate(Workbook workbook, String sheetName) {
        /* Open sheet */
        Sheet sheet = workbook.getSheet(sheetName);
        return sheet == null ? workbook.createSheet(sheetName) : sheet;
    }

    /**
     * Check if the sheet is present
     * @param workbook The {@code Workbook} where there is the sheet
     * @param sheetName The sheet name in the workbook
     * @return {@code true} if is present
     */
    public static Boolean isPresent(Workbook workbook, String sheetName) {
        return workbook.getSheet(sheetName) != null;
    }

    /**
     * Check if the sheet is present
     * @param workbook The {@code Workbook} where there is the sheet
     * @param position The index in the workbook
     * @return {@code true} if is present
     */
    public static Boolean isPresent(Workbook workbook, Integer position) {
        return workbook.getSheetAt(position) != null;
    }

    /**
     * Check if the sheet is empty
     * @param workbook The {@code Workbook} where there is the sheet
     * @param sheetName The sheet name in the workbook
     * @return {@code true} if is empty
     */
    public static Boolean isNull(Workbook workbook, String sheetName) {
        return workbook.getSheet(sheetName) == null;
    }

    /**
     * Check if the sheet is empty
     * @param workbook The {@code Workbook} where there is the sheet
     * @param position The index in the workbook
     * @return {@code true} if is empty
     */
    public static Boolean isNull(Workbook workbook, Integer position) {
        return workbook.getSheetAt(position) == null;
    }
}
