package tools.implementations;

import enums.ExcelExtension;
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

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWorkbookUtilsImpl implements ExcelWorkbookUtils {

    @Override
    public Workbook open(FileInputStream fileInputStream, String extension) throws ExtensionNotValidException, IOException, OpenWorkbookException {

        /* Check the extension */
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        if(!excelUtils.isValidExcelExtension(extension)) {
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

    @Override
    public Workbook create() {
        return create(ExcelExtension.XLSX);
    }

    @Override
    public Workbook create(String extension) throws ExtensionNotValidException {
        ExcelUtils excelUtils = new ExcelUtilsImpl();
        if(!excelUtils.isValidExcelExtension(extension)) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }
        return create(ExcelExtension.getExcelExtension(extension));
    }

    @Override
    public Workbook create(ExcelExtension extension) {
        Workbook workbook = null;
        switch (extension) {
            case XLS -> workbook = new HSSFWorkbook();
            case XLSX -> workbook = new XSSFWorkbook();
        }
        return workbook;
    }

    @Override
    public void close(Workbook workbook) throws IOException {
        workbook.close();
    }

    @Override
    public void close(Workbook workbook, FileInputStream fileInputStream) throws IOException {
        workbook.close();
        fileInputStream.close();
    }

    @Override
    public void close(Workbook workbook, FileOutputStream fileOutputStream) throws IOException {
        workbook.close();
        fileOutputStream.close();
    }
}
