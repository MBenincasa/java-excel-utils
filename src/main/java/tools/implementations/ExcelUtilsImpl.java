package tools.implementations;

import enums.ExcelExtension;
import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import tools.interfaces.ExcelSheetUtils;
import tools.interfaces.ExcelUtils;
import tools.interfaces.ExcelWorkbookUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelUtilsImpl implements ExcelUtils {

    @Override
    public Integer countAllRows(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        return countAllRows(file, true, null);
    }

    @Override
    public Integer countAllRows(File file, Boolean alsoEmptyRows) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        return countAllRows(file, alsoEmptyRows, null);
    }

    @Override
    public Integer countAllRows(File file, Boolean alsoEmptyRows, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {

        /* Check extension */
        String extension = checkExtension(file.getName());

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
                ? sheet.getPhysicalNumberOfRows()
                : countOnlyRowsNotEmpty(sheet);

        /* Close file */
        excelWorkbookUtils.close(workbook, fileInputStream);

        return numRows;
    }

    @Override
    public Boolean isValidExcelExtension(String extension) {
        return extension.equalsIgnoreCase(ExcelExtension.XLS.getExt()) || extension.equalsIgnoreCase(ExcelExtension.XLSX.getExt());
    }

    private int countOnlyRowsNotEmpty(Sheet sheet) {
        int numRows = sheet.getPhysicalNumberOfRows();
        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            boolean isEmptyRow = true;

            if (row != null) {
                for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                    Cell cell = row.getCell(j);

                    if (cell.getStringCellValue() != null && !cell.getStringCellValue().isEmpty()) {
                        isEmptyRow = false;
                        break;
                    }
                }
            }

            if(isEmptyRow) {
                numRows--;
            }
        }

        return numRows;
    }

    private String checkExtension(String filename) throws ExtensionNotValidException {
        String extension = FilenameUtils.getExtension(filename);
        if(!isValidExcelExtension(extension)) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }
        return extension;
    }

}
