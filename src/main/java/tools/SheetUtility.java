package tools;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetAlreadyExistsException;
import exceptions.SheetNotFoundException;
import model.ExcelSheet;
import model.ExcelWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

/**
 * {@code SheetUtility} is the static class with implementations of some methods for working with Sheets
 * @deprecated since version 0.3.0. View here {@link model.ExcelSheet}
 * @see model.ExcelSheet
 * @author Mirko Benincasa
 * @since 0.2.0
 */
@Deprecated
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
        Integer totalSheets = excelWorkbook.length();

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

        /* Close file */
        excelWorkbook.close();

        return excelWorkbook.getSheets().stream().map(ExcelSheet::getName).toList();
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
        List<ExcelSheet> excelSheets = excelWorkbook.getSheets();
        Optional<ExcelSheet> excelSheet = excelSheets.stream().filter(s -> s.getName().equals(sheetName)).findFirst();

        if (excelSheet.isEmpty())
            throw new SheetNotFoundException("No sheet was found");

        /* Close file */
        excelWorkbook.close();

        return excelSheet.get().getIndex();
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
        List<ExcelSheet> excelSheets = excelWorkbook.getSheets();
        Optional<ExcelSheet> excelSheet = excelSheets.stream().filter(s -> Objects.equals(s.getIndex(), position)).findFirst();

        if (excelSheet.isEmpty())
            throw new SheetNotFoundException("No sheet was found");

        /* Close file */
        excelWorkbook.close();

        return excelSheet.get().getName();
    }

    /**
     * Create a sheet in a workbook
     * @param file Excel file
     * @return The new sheet that was created
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error occurs
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     * @throws SheetAlreadyExistsException If you try to insert a sheet with a name that already exists
     */
    public static Sheet create(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetAlreadyExistsException {
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
     * @throws SheetAlreadyExistsException If you try to insert a sheet with a name that already exists
     */
    public static Sheet create(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetAlreadyExistsException {
        /* Check extension */
        String extension = ExcelUtility.checkExcelExtension(file.getName());

        /* Open file excel */
        FileInputStream fileInputStream = new FileInputStream(file);
        ExcelWorkbook excelWorkbook = ExcelWorkbook.open(fileInputStream, extension);
        ExcelSheet excelSheet = ExcelSheet.create(excelWorkbook, sheetName);

        return excelSheet.getSheet();
    }

    /**
     * Create a sheet in a workbook
     * @param workbook The {@code Workbook} where to add the new sheet
     * @return The new sheet that was created
     * @throws SheetAlreadyExistsException If you try to insert a sheet with a name that already exists
     */
    public static Sheet create(Workbook workbook) throws SheetAlreadyExistsException {
        return create(workbook, null);
    }

    /**
     * Create a sheet in a workbook
     * @param workbook The {@code Workbook} where to add the new sheet
     * @param sheetName The name of the sheet to add
     * @return The new sheet that was created
     * @throws SheetAlreadyExistsException If you try to insert a sheet with a name that already exists
     */
    public static Sheet create(Workbook workbook, String sheetName) throws SheetAlreadyExistsException {
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(workbook);
        ExcelSheet excelSheet = ExcelSheet.create(excelWorkbook, sheetName);
        return excelSheet.getSheet();
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
        ExcelSheet excelSheet = excelWorkbook.getSheet(sheetName);

        /* Close workbook */
        excelWorkbook.close();

        return excelSheet.getSheet();
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
        ExcelSheet excelSheet = excelWorkbook.getSheet(position);

        /* Close workbook */
        excelWorkbook.close();

        return excelSheet.getSheet();
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
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(workbook);
        ExcelSheet excelSheet = excelWorkbook.getSheet(sheetName);
        return excelSheet.getSheet();
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
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(workbook);
        ExcelSheet excelSheet = excelWorkbook.getSheet(position);
        return excelSheet.getSheet();
    }

    /**
     * Gets the sheet in the workbook. If it doesn't find it, it creates it.
     * @param workbook The {@code Workbook} where there is the sheet
     * @param sheetName The sheet name in the workbook
     * @return The sheet in the workbook or a new one
     */
    public static Sheet getOrCreate(Workbook workbook, String sheetName) {
        /* Open sheet */
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(workbook);
        ExcelSheet excelSheet = excelWorkbook.getSheetOrCreate(sheetName);
        return excelSheet.getSheet();
    }

    /**
     * Check if the sheet is present
     * @param workbook The {@code Workbook} where there is the sheet
     * @param sheetName The sheet name in the workbook
     * @return {@code true} if is present
     */
    public static Boolean isPresent(Workbook workbook, String sheetName) {
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(workbook);
        return excelWorkbook.isSheetPresent(sheetName);
    }

    /**
     * Check if the sheet is present
     * @param workbook The {@code Workbook} where there is the sheet
     * @param position The index in the workbook
     * @return {@code true} if is present
     */
    public static Boolean isPresent(Workbook workbook, Integer position) {
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(workbook);
        return excelWorkbook.isSheetPresent(position);
    }

    /**
     * Check if the sheet is empty
     * @param workbook The {@code Workbook} where there is the sheet
     * @param sheetName The sheet name in the workbook
     * @return {@code true} if is empty
     */
    public static Boolean isNull(Workbook workbook, String sheetName) {
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(workbook);
        return excelWorkbook.isSheetNull(sheetName);
    }

    /**
     * Check if the sheet is empty
     * @param workbook The {@code Workbook} where there is the sheet
     * @param position The index in the workbook
     * @return {@code true} if is empty
     */
    public static Boolean isNull(Workbook workbook, Integer position) {
        ExcelWorkbook excelWorkbook = new ExcelWorkbook(workbook);
        return excelWorkbook.isSheetNull(position);
    }
}
