package tools.utils;

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

public class ExcelUtilsImpl implements ExcelUtils {

    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";

    @Override
    public Integer countAllRows(File file, Boolean alsoEmptyRows) throws Exception {

        /* Check extension */
        String extension = FilenameUtils.getExtension(file.getName());
        if(!extension.equalsIgnoreCase(XLS) && !extension.equalsIgnoreCase(XLSX)) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }

        /* Open file excel */
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = openWorkbook(fileInputStream, extension);
        Sheet sheet = workbook.getSheetAt(0);

        /* Count all rows */
        int numRows = alsoEmptyRows
                ? sheet.getPhysicalNumberOfRows()
                : removeEmptyRows(sheet);

        /* Close file */
        fileInputStream.close();
        workbook.close();

        return numRows;
    }

    private Workbook openWorkbook(FileInputStream fileInputStream, String extension) throws Exception {
        Workbook workbook;
        switch (extension) {
            case XLS -> workbook = new HSSFWorkbook(fileInputStream);
            case XLSX -> workbook = new XSSFWorkbook(fileInputStream);
            default -> throw new Exception();
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
