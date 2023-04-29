package io.github.mbenincasa.javaexcelutils.tools;

import io.github.mbenincasa.javaexcelutils.enums.Extension;
import io.github.mbenincasa.javaexcelutils.exceptions.ExtensionNotValidException;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.util.CellReference;

/**
 * {@code ExcelUtility} is the static class with the implementations of some utilities on Excel files
 * @author Mirko Benincasa
 * @since 0.2.0
 */
public class ExcelUtility {

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

    /**
     * Returns the cell name
     * @param row row index
     * @param col column index
     * @return cell name
     * @since 0.4.2
     */
    public static String getCellName(int row, int col) {
        String colName = CellReference.convertNumToColString(col);
        return colName + (row + 1);
    }

    /**
     * Return an array containing column and row indexes
     * @param cellName cell name
     * @return an array containing column and row indexes
     * @since 0.4.2
     */
    public static int[] getCellIndexes(String cellName) {
        CellReference cellRef = new CellReference(cellName);
        int row = cellRef.getRow();
        int col = cellRef.getCol();
        return new int[]{row, col};
    }
}
