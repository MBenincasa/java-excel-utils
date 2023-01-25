package tools.interfaces;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * The {@code ExcelSheetUtils} interface groups methods that work with sheets
 * @deprecated since version 0.2.0. View here {@link tools.SheetUtility}
 * @see tools.SheetUtility
 * @author Mirko Benincasa
 * @since 0.1.0
 */
@Deprecated
public interface ExcelSheetUtils {

    /**
     * Count how many sheets there are
     * @param file Excel file
     * @return The number of sheets present
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    Integer countAll(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    /**
     * Returns the name of all sheets in the workbook
     * @param file Excel file
     * @return A list with the name of all sheets
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    List<String> getAllNames(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException;

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
    Integer getIndex(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

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
    String getNameByIndex(File file, Integer position) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    /**
     * Create a sheet in a workbook
     * @param file Excel file
     * @return The new sheet that was created
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    Sheet create(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    /**
     * Create a sheet in a workbook
     * @param file Excel file
     * @param sheetName The name of the sheet to add
     * @return The new sheet that was created
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    Sheet create(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    /**
     * Create a sheet in a workbook
     * @param workbook The {@code Workbook} where to add the new sheet
     * @return The new sheet that was created
     */
    Sheet create(Workbook workbook);

    /**
     * Create a sheet in a workbook
     * @param workbook The {@code Workbook} where to add the new sheet
     * @param sheetName The name of the sheet to add
     * @return The new sheet that was created
     */
    Sheet create(Workbook workbook, String sheetName);

    /**
     * Opens the sheet of the Excel file
     * @param file Excel file
     * @return The sheet in the workbook
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    Sheet open(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    /**
     * Opens the sheet of the Excel file
     * @param file Excel file
     * @param sheetName The sheet name in the workbook
     * @return The sheet in the workbook
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    Sheet open(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    /**
     * Opens the sheet of the Excel file
     * @param file Excel file
     * @param position The index in the workbook
     * @return The sheet in the workbook
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    Sheet open(File file, Integer position) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    /**
     * Opens the sheet in the workbook.
     * @param workbook The {@code Workbook} where there is the sheet
     * @return The sheet in the workbook in first position
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    Sheet open(Workbook workbook) throws SheetNotFoundException;

    /**
     * Opens the sheet in the workbook.
     * @param workbook The {@code Workbook} where there is the sheet
     * @param sheetName The sheet name in the workbook
     * @return The sheet in the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    Sheet open(Workbook workbook, String sheetName) throws SheetNotFoundException;

    /**
     * Opens the sheet in the workbook.
     * @param workbook The {@code Workbook} where there is the sheet
     * @param position The index in the workbook
     * @return The sheet in the workbook
     * @throws SheetNotFoundException If the sheet to open is not found
     */
    Sheet open(Workbook workbook, Integer position) throws SheetNotFoundException;

    /**
     * Opens the sheet in the workbook. If it doesn't find it, it creates it.
     * @param workbook The {@code Workbook} where there is the sheet
     * @param sheetName The sheet name in the workbook
     * @return The sheet in the workbook or a new one
     */
    Sheet openOrCreate(Workbook workbook, String sheetName);

    /**
     * Check if the sheet is present
     * @param workbook The {@code Workbook} where there is the sheet
     * @param sheetName The sheet name in the workbook
     * @return {@code true} if is present
     */
    Boolean isPresent(Workbook workbook, String sheetName);

    /**
     * Check if the sheet is present
     * @param workbook The {@code Workbook} where there is the sheet
     * @param position The index in the workbook
     * @return {@code true} if is present
     */
    Boolean isPresent(Workbook workbook, Integer position);
}
