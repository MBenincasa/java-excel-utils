package tools.implementations;

import com.opencsv.CSVWriter;
import enums.Extension;
import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import tools.interfaces.ExcelUtils;
import tools.interfaces.ExcelWorkbookUtils;

import java.io.*;

/**
 * {@code ExcelWorkbookUtilsImpl} is the standard implementation class of {@code ExcelWorkbookUtils}
 * @author Mirko Benincasa
 * @since 0.1.0
 */
public class ExcelWorkbookUtilsImpl implements ExcelWorkbookUtils {

    /**
     * {@inheritDoc}
     * @param file {@inheritDoc}
     * @return {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws OpenWorkbookException {@inheritDoc}
     */
    @Override
    public Workbook open(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Check extension */
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        String extension = excelUtils.checkExcelExtension(file.getName());

        /* Open file input stream */
        FileInputStream fileInputStream = new FileInputStream(file);
        return open(fileInputStream, extension);
    }

    /**
     * {@inheritDoc}
     * @param fileInputStream {@inheritDoc}
     * @param extension {@inheritDoc}
     * @return {@inheritDoc}
     * @throws ExtensionNotValidException {@inheritDoc}
     * @throws IOException {@inheritDoc}
     * @throws OpenWorkbookException {@inheritDoc}
     */
    @Override
    public Workbook open(FileInputStream fileInputStream, String extension) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Check the extension */
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        if (!excelUtils.isValidExcelExtension(extension)) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }

        /* Open workbook */
        try {
            return new XSSFWorkbook(fileInputStream);
        } catch (OfficeXmlFileException | OLE2NotOfficeXmlFileException e) {
            try {
                return new HSSFWorkbook(fileInputStream);
            } catch (NotOLE2FileException ex) {
                throw new OpenWorkbookException("The workbook could not be opened", ex);
            }
        }
    }

    /**
     * {@inheritDoc}<p>
     * If not specified the XLSX extension will be used
     * @return {@inheritDoc}
     */
    @Override
    public Workbook create() {
        return create(Extension.XLSX);
    }

    /**
     * {@inheritDoc}
     * @param extension {@inheritDoc}
     * @return {@inheritDoc}
     * @throws {@inheritDoc}
     */
    @Override
    public Workbook create(String extension) throws ExtensionNotValidException {
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        if (!excelUtils.isValidExcelExtension(extension)) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }
        return create(Extension.getExcelExtension(extension));
    }

    /**
     * {@inheritDoc}
     * @param extension {@inheritDoc}
     * @return {@inheritDoc}
     */
    @Override
    public Workbook create(Extension extension) {
        Workbook workbook = null;
        switch (extension) {
            case XLS -> workbook = new HSSFWorkbook();
            case XLSX -> workbook = new XSSFWorkbook();
        }
        return workbook;
    }

    /**
     * {@inheritDoc}
     * @param workbook {@inheritDoc}
     * @throws IOException {@inheritDoc}
     */
    @Override
    public void close(Workbook workbook) throws IOException {
        workbook.close();
    }

    /**
     * {@inheritDoc}
     * @param workbook {@inheritDoc}
     * @param fileInputStream {@inheritDoc}
     * @throws IOException {@inheritDoc}
     */
    @Override
    public void close(Workbook workbook, FileInputStream fileInputStream) throws IOException {
        workbook.close();
        fileInputStream.close();
    }

    /**
     * {@inheritDoc}
     * @param workbook {@inheritDoc}
     * @param fileOutputStream {@inheritDoc}
     * @throws IOException {@inheritDoc}
     */
    @Override
    public void close(Workbook workbook, FileOutputStream fileOutputStream) throws IOException {
        workbook.close();
        fileOutputStream.close();
    }

    /**
     * {@inheritDoc}
     * @param workbook {@inheritDoc}
     * @param fileOutputStream {@inheritDoc}
     * @param fileInputStream {@inheritDoc}
     * @throws IOException {@inheritDoc}
     */
    @Override
    public void close(Workbook workbook, FileOutputStream fileOutputStream, FileInputStream fileInputStream) throws IOException {
        workbook.close();
        fileInputStream.close();
        fileOutputStream.close();
    }

    /**
     * {@inheritDoc}
     * @param workbook {@inheritDoc}
     * @param fileInputStream {@inheritDoc}
     * @param writer {@inheritDoc}
     * @throws IOException {@inheritDoc}
     */
    @Override
    public void close(Workbook workbook, FileInputStream fileInputStream, CSVWriter writer) throws IOException {
        workbook.close();
        fileInputStream.close();
        writer.close();
    }
}
