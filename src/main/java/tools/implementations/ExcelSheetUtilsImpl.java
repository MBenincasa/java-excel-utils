package tools.implementations;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import tools.interfaces.ExcelSheetUtils;
import tools.interfaces.ExcelUtils;
import tools.interfaces.ExcelWorkbookUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

public class ExcelSheetUtilsImpl implements ExcelSheetUtils {

    @Override
    public Integer countAllSheets(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {

        /* Check extension */
        String extension = checkExtension(file.getName());

        /* Open file excel */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = excelWorkbookUtils.openWorkbook(fileInputStream, extension);

        Integer totalSheets = workbook.getNumberOfSheets();

        /* Close file */
        closeFile(workbook, fileInputStream);

        return totalSheets;
    }

    @Override
    public List<String> getAllSheetNames(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {

        /* Check extension */
        String extension = checkExtension(file.getName());

        /* Open file excel */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = excelWorkbookUtils.openWorkbook(fileInputStream, extension);

        /* Iterate all the sheets */
        Iterator<Sheet> sheetIterator = workbook.iterator();
        List<String> sheetNames = new LinkedList<>();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            sheetNames.add(sheet.getSheetName());
        }

        /* Close file */
        closeFile(workbook, fileInputStream);

        return sheetNames;
    }

    @Override
    public Integer getSheetIndex(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {

        /* Check extension */
        String extension = checkExtension(file.getName());

        /* Open file excel */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = excelWorkbookUtils.openWorkbook(fileInputStream, extension);

        int sheetIndex = workbook.getSheetIndex(sheetName);

        /* Close file */
        closeFile(workbook, fileInputStream);

        if(sheetIndex < 0) {
            throw new SheetNotFoundException("No sheet was found");
        }
        return sheetIndex;
    }

    @Override
    public String getSheetNameAtPosition(File file, Integer position) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {

        /* Check extension */
        String extension = checkExtension(file.getName());

        /* Open file excel */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = excelWorkbookUtils.openWorkbook(fileInputStream, extension);

        String sheetName;
        try {
            sheetName = workbook.getSheetName(position);
        } catch (IllegalArgumentException e) {
            throw new SheetNotFoundException("Sheet index is out of range");
        }

        /* Close file */
        closeFile(workbook, fileInputStream);

        return sheetName;
    }

    private String checkExtension(String filename) throws ExtensionNotValidException {
        String extension = FilenameUtils.getExtension(filename);
        ExcelUtils excelUtils = new ExcelUtilsImpl();

        if(!excelUtils.isValidExcelExtension(extension)) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }
        return extension;
    }

    private void closeFile(Workbook workbook, FileInputStream fileInputStream) throws IOException {
        fileInputStream.close();
        workbook.close();
    }
}
