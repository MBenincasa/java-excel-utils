package tools;

import enums.Extension;
import exceptions.ExtensionNotValidException;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelUtilsImpl implements ExcelUtils {

    @Override
    public Integer countAllRows(File file) throws ExtensionNotValidException, IOException {
        return countAllRows(file, true, null);
    }

    @Override
    public Integer countAllRows(File file, Boolean alsoEmptyRows) throws ExtensionNotValidException, IOException {
        return countAllRows(file, alsoEmptyRows, null);
    }

    @Override
    public Integer countAllRows(File file, Boolean alsoEmptyRows, String sheetName) throws ExtensionNotValidException, IOException {

        /* Check extension */
        String extension = FilenameUtils.getExtension(file.getName());
        if(!extension.equalsIgnoreCase(Extension.XLS.getExt()) && !extension.equalsIgnoreCase(Extension.XLSX.getExt())) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }

        /* Open file excel */
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = openWorkbook(fileInputStream, extension);
        Sheet sheet = (sheetName == null || sheetName.isEmpty())
                ? workbook.getSheetAt(0)
                : workbook.getSheet(sheetName);

        /* Count all rows */
        int numRows = alsoEmptyRows
                ? sheet.getPhysicalNumberOfRows()
                : removeEmptyRows(sheet);

        /* Close file */
        fileInputStream.close();
        workbook.close();

        return numRows;
    }

    private Workbook openWorkbook(FileInputStream fileInputStream, String extension) throws ExtensionNotValidException, IOException {
        Workbook workbook;
        switch (extension) {
            case "xls" -> workbook = new HSSFWorkbook(fileInputStream);
            case "xlsx" -> workbook = new XSSFWorkbook(fileInputStream);
            default -> throw new ExtensionNotValidException();
        }
        return workbook;
    }

    private int removeEmptyRows(Sheet sheet) {
        int numRows = sheet.getPhysicalNumberOfRows();
        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            boolean isEmptyRow = true;

            for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                Cell cell = row.getCell(j);

                if (cell.getStringCellValue() != null && !cell.getStringCellValue().isEmpty()) {
                    isEmptyRow = false;
                    break;
                }
            }

            if(isEmptyRow) {
                numRows--;
            }
        }

        return numRows;
    }
}
