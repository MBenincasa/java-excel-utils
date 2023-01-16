package tools.implementations;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;
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
    public Integer countAll(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Check extension */
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        String extension = excelUtils.checkExcelExtension(file.getName());

        /* Open file excel */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = excelWorkbookUtils.open(fileInputStream, extension);

        Integer totalSheets = workbook.getNumberOfSheets();

        /* Close file */
        excelWorkbookUtils.close(workbook, fileInputStream);

        return totalSheets;
    }

    @Override
    public List<String> getAllNames(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Check extension */
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        String extension = excelUtils.checkExcelExtension(file.getName());

        /* Open file excel */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = excelWorkbookUtils.open(fileInputStream, extension);

        /* Iterate all the sheets */
        Iterator<Sheet> sheetIterator = workbook.iterator();
        List<String> sheetNames = new LinkedList<>();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            sheetNames.add(sheet.getSheetName());
        }

        /* Close file */
        excelWorkbookUtils.close(workbook, fileInputStream);

        return sheetNames;
    }

    @Override
    public Integer getIndex(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        /* Check extension */
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        String extension = excelUtils.checkExcelExtension(file.getName());

        /* Open file excel */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = excelWorkbookUtils.open(fileInputStream, extension);

        int sheetIndex = workbook.getSheetIndex(sheetName);

        /* Close file */
        excelWorkbookUtils.close(workbook, fileInputStream);

        if (sheetIndex < 0) {
            throw new SheetNotFoundException("No sheet was found");
        }
        return sheetIndex;
    }

    @Override
    public String getNameByIndex(File file, Integer position) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        /* Check extension */
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        String extension = excelUtils.checkExcelExtension(file.getName());

        /* Open file excel */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = excelWorkbookUtils.open(fileInputStream, extension);

        String sheetName;
        try {
            sheetName = workbook.getSheetName(position);
        } catch (IllegalArgumentException e) {
            throw new SheetNotFoundException("Sheet index is out of range");
        }

        /* Close file */
        excelWorkbookUtils.close(workbook, fileInputStream);

        return sheetName;
    }

    @Override
    public Sheet create(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        return create(file, null);
    }

    @Override
    public Sheet create(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Check extension */
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        String extension = excelUtils.checkExcelExtension(file.getName());

        /* Open file excel */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = excelWorkbookUtils.open(fileInputStream, extension);

        /* Create sheet */
        return sheetName == null ? workbook.createSheet() : workbook.createSheet(sheetName);
    }

    @Override
    public Sheet create(Workbook workbook) {
        return create(workbook, null);
    }

    @Override
    public Sheet create(Workbook workbook, String sheetName) {
        return sheetName == null ? workbook.createSheet() : workbook.createSheet(sheetName);
    }

    @Override
    public Sheet open(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        return open(file, 0);
    }

    @Override
    public Sheet open(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        /* Check extension */
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        String extension = excelUtils.checkExcelExtension(file.getName());

        /* Open file excel */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = excelWorkbookUtils.open(fileInputStream, extension);

        /* Open sheet */
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null)
            throw new SheetNotFoundException();
        return sheet;
    }

    @Override
    public Sheet open(File file, Integer position) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        /* Check extension */
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        String extension = excelUtils.checkExcelExtension(file.getName());

        /* Open file excel */
        ExcelWorkbookUtils excelWorkbookUtils = new ExcelWorkbookUtilsImpl();
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = excelWorkbookUtils.open(fileInputStream, extension);

        /* Open sheet */
        Sheet sheet = workbook.getSheetAt(position);
        if (sheet == null)
            throw new SheetNotFoundException();
        return sheet;
    }

    @Override
    public Sheet open(Workbook workbook) throws SheetNotFoundException {
        return open(workbook, 0);
    }

    @Override
    public Sheet open(Workbook workbook, String sheetName) throws SheetNotFoundException {
        /* Open sheet */
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null)
            throw new SheetNotFoundException();
        return sheet;
    }

    @Override
    public Sheet open(Workbook workbook, Integer position) throws SheetNotFoundException {
        /* Open sheet */
        Sheet sheet = workbook.getSheetAt(position);
        if (sheet == null)
            throw new SheetNotFoundException();
        return sheet;
    }

    @Override
    public Sheet openOrCreate(Workbook workbook, String sheetName) {
        /* Open sheet */
        Sheet sheet = workbook.getSheet(sheetName);
        return sheet == null ? workbook.createSheet(sheetName) : sheet;
    }

    @Override
    public Boolean isPresent(Workbook workbook, String sheetName) {
        return workbook.getSheet(sheetName) != null;
    }

    @Override
    public Boolean isPresent(Workbook workbook, Integer position) {
        return workbook.getSheetAt(position) != null;
    }
}
