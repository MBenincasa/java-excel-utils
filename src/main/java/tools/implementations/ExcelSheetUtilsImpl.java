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

/**
 * {@code ExcelSheetUtilsImpl} is the standard implementation class of {@code ExcelSheetUtils}
 * @deprecated since version 0.2.0. View here {@link tools.SheetUtility}
 * @see tools.SheetUtility
 * @author Mirko Benincasa
 * @since 0.1.0
 */
@Deprecated
public class ExcelSheetUtilsImpl implements ExcelSheetUtils {

    /**
     * {@inheritDoc}
     * @param file {@inheritDoc}
     * @return {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws OpenWorkbookException {@inheritDoc}
     */
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

    /**
     * {@inheritDoc}
     * @param file {@inheritDoc}
     * @return {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws OpenWorkbookException {@inheritDoc}
     */
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

    /**
     * {@inheritDoc}
     * @param file {@inheritDoc}
     * @param sheetName {@inheritDoc}
     * @return {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws OpenWorkbookException {@inheritDoc}
     * @throws SheetNotFoundException {@inheritDoc}
     */
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

    /**
     * {@inheritDoc}
     * @param file {@inheritDoc}
     * @param position {@inheritDoc}
     * @return {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws OpenWorkbookException {@inheritDoc}
     * @throws SheetNotFoundException {@inheritDoc}
     */
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

    /**
     * {@inheritDoc}
     * @param file {@inheritDoc}
     * @return {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws OpenWorkbookException {@inheritDoc}
     */
    @Override
    public Sheet create(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        return create(file, null);
    }

    /**
     * {@inheritDoc}
     * @param file {@inheritDoc}
     * @param sheetName {@inheritDoc}
     * @return {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws OpenWorkbookException {@inheritDoc}
     */
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

    /**
     * {@inheritDoc}
     * @param workbook {@inheritDoc}
     * @return {@inheritDoc}
     */
    @Override
    public Sheet create(Workbook workbook) {
        return create(workbook, null);
    }

    /**
     * {@inheritDoc}
     * @param workbook {@inheritDoc}
     * @param sheetName {@inheritDoc}
     * @return {@inheritDoc}
     */
    @Override
    public Sheet create(Workbook workbook, String sheetName) {
        return sheetName == null ? workbook.createSheet() : workbook.createSheet(sheetName);
    }

    /**
     * {@inheritDoc}<p>
     * If not specified, the first sheet will be opened
     * @param file {@inheritDoc}
     * @return {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws OpenWorkbookException {@inheritDoc}
     * @throws SheetNotFoundException {@inheritDoc}
     */
    @Override
    public Sheet open(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException {
        return open(file, 0);
    }

    /**
     * {@inheritDoc}
     * @param file {@inheritDoc}
     * @param sheetName {@inheritDoc}
     * @return {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws OpenWorkbookException {@inheritDoc}
     * @throws SheetNotFoundException {@inheritDoc}
     */
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

    /**
     * {@inheritDoc}
     * @param file {@inheritDoc}
     * @param position {@inheritDoc}
     * @return {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws OpenWorkbookException {@inheritDoc}
     * @throws SheetNotFoundException {@inheritDoc}
     */
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

    /**
     * {@inheritDoc}<p>
     * If not specified, the first sheet will be opened
     * @param workbook {@inheritDoc}
     * @return {@inheritDoc}
     * @throws SheetNotFoundException {@inheritDoc}
     */
    @Override
    public Sheet open(Workbook workbook) throws SheetNotFoundException {
        return open(workbook, 0);
    }

    /**
     * {@inheritDoc}
     * @param workbook {@inheritDoc}
     * @param sheetName {@inheritDoc}
     * @return {@inheritDoc}
     * @throws SheetNotFoundException {@inheritDoc}
     */
    @Override
    public Sheet open(Workbook workbook, String sheetName) throws SheetNotFoundException {
        /* Open sheet */
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null)
            throw new SheetNotFoundException();
        return sheet;
    }

    /**
     * {@inheritDoc}
     * @param workbook {@inheritDoc}
     * @param position {@inheritDoc}
     * @return {@inheritDoc}
     * @throws SheetNotFoundException {@inheritDoc}
     */
    @Override
    public Sheet open(Workbook workbook, Integer position) throws SheetNotFoundException {
        /* Open sheet */
        Sheet sheet = workbook.getSheetAt(position);
        if (sheet == null)
            throw new SheetNotFoundException();
        return sheet;
    }

    /**
     * {@inheritDoc}
     * @param workbook {@inheritDoc}
     * @param sheetName {@inheritDoc}
     * @return {@inheritDoc}
     */
    @Override
    public Sheet openOrCreate(Workbook workbook, String sheetName) {
        /* Open sheet */
        Sheet sheet = workbook.getSheet(sheetName);
        return sheet == null ? workbook.createSheet(sheetName) : sheet;
    }

    /**
     * {@inheritDoc}
     * @param workbook {@inheritDoc}
     * @param sheetName {@inheritDoc}
     * @return {@inheritDoc}
     */
    @Override
    public Boolean isPresent(Workbook workbook, String sheetName) {
        return workbook.getSheet(sheetName) != null;
    }

    /**
     * {@inheritDoc}
     * @param workbook {@inheritDoc}
     * @param position {@inheritDoc}
     * @return {@inheritDoc}
     */
    @Override
    public Boolean isPresent(Workbook workbook, Integer position) {
        return workbook.getSheetAt(position) != null;
    }
}
