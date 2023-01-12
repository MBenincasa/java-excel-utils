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
import java.util.LinkedList;
import java.util.List;

public class ExcelUtilsImpl implements ExcelUtils {

    @Override
    public List<Integer> countAllRowsOfAllSheets(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        return countAllRowsOfAllSheets(file, true);
    }

    @Override
    public List<Integer> countAllRowsOfAllSheets(File file, Boolean alsoEmptyRows) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Check extension */
        String extension = this.checkExtension(file.getName());

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

            values.add(this.countOnlyRowsNotEmpty(sheet));
        }

        /* Close file */
        excelWorkbookUtils.close(workbook, fileInputStream);

        return values;
    }

    @Override
    public Integer countAllRows(File file, String sheetName) throws OpenWorkbookException, SheetNotFoundException, ExtensionNotValidException, IOException {
        return countAllRows(file, sheetName, true);
    }

    @Override
    public Integer countAllRows(File file, String sheetName, Boolean alsoEmptyRows) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        /* Check extension */
        String extension = this.checkExtension(file.getName());

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
                : this.countOnlyRowsNotEmpty(sheet);

        /* Close file */
        excelWorkbookUtils.close(workbook, fileInputStream);

        return numRows;
    }

    @Override
    public Boolean isValidExcelExtension(String extension) {
        return extension.equalsIgnoreCase(ExcelExtension.XLS.getExt()) || extension.equalsIgnoreCase(ExcelExtension.XLSX.getExt());
    }

    @Override
    public String checkExtension(String filename) throws ExtensionNotValidException {
        String extension = FilenameUtils.getExtension(filename);
        if (!isValidExcelExtension(extension)) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }
        return extension;
    }

    private int countOnlyRowsNotEmpty(Sheet sheet) {
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
