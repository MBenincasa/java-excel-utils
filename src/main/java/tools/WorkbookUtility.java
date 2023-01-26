package tools;

import com.opencsv.CSVReader;
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

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * {@code WorkbookUtility} is the static class with implementations of some methods for working with Workbooks
 * @author Mirko Benincasa
 * @since 0.2.0
 */
public class WorkbookUtility {

    /**
     * Opens the workbook
     * @param file An Excel file
     * @return The workbook that was opened
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    public static Workbook open(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Check extension */
        String extension = ExcelUtility.checkExcelExtension(file.getName());

        /* Open file input stream */
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = open(fileInputStream, extension);

        /* Close the stream before return */
        fileInputStream.close();
        return workbook;
    }

    /**
     * Opens the workbook
     * @param fileInputStream The {@code FileInputStream} of the Excel file
     * @param extension The file's extension
     * @return The workbook that was opened
     * @throws ExtensionNotValidException If the input file extension does not belong to an Excel file
     * @throws IOException If an I/O error has occurred
     * @throws OpenWorkbookException If an error occurred while opening the workbook
     */
    public static Workbook open(FileInputStream fileInputStream, String extension) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Check the extension */
        if (!ExcelUtility.isValidExcelExtension(extension)) {
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
     * Create a new workbook<p>
     * If not specified the XLSX extension will be used
     * @return A workbook
     */
    public static Workbook create() {
        return create(Extension.XLSX);
    }

    /**
     * Create a new workbook
     * @param extension The extension of the file. Provide the extension of an Excel file
     * @return A workbook
     * @throws ExtensionNotValidException If the extension does not belong to an Excel file
     */
    public static Workbook create(String extension) throws ExtensionNotValidException {
        if (!ExcelUtility.isValidExcelExtension(extension)) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }
        return create(Extension.getExcelExtension(extension));
    }

    /**
     * Create a new workbook
     * @param extension The extension of the file. Select an extension with {@code type} EXCEL
     * @return A workbook
     */
    public static Workbook create(Extension extension) {
        Workbook workbook = null;
        switch (extension) {
            case XLS -> workbook = new HSSFWorkbook();
            case XLSX -> workbook = new XSSFWorkbook();
        }
        return workbook;
    }

    /**
     * Close a workbook
     * @param workbook The {@code Workbook} to close
     * @throws IOException If an I/O error has occurred
     */
    public static void close(Workbook workbook) throws IOException {
        workbook.close();
    }

    /**
     * Close a workbook
     * @param workbook The {@code Workbook} to close
     * @param fileInputStream The {@code FileInputStream} to close
     * @throws IOException If an I/O error has occurred
     */
    public static void close(Workbook workbook, FileInputStream fileInputStream) throws IOException {
        workbook.close();
        fileInputStream.close();
    }

    /**
     * Close a workbook
     * @param workbook The {@code Workbook} to close
     * @param fileOutputStream The {@code FileOutputStream} to close
     * @throws IOException If an I/O error has occurred
     */
    public static void close(Workbook workbook, FileOutputStream fileOutputStream) throws IOException {
        workbook.close();
        fileOutputStream.close();
    }

    /**
     * Close a workbook
     * @param workbook The {@code Workbook} to close
     * @param fileOutputStream The {@code FileOutputStream} to close
     * @param fileInputStream The {@code FileInputStream} to close
     * @throws IOException If an I/O error has occurred
     */
    public static void close(Workbook workbook, FileOutputStream fileOutputStream, FileInputStream fileInputStream) throws IOException {
        workbook.close();
        fileInputStream.close();
        fileOutputStream.close();
    }

    /**
     * Close a workbook
     * @param workbook The {@code Workbook} to close
     * @param writer The {@code CSVWriter} to close
     * @throws IOException If an I/O error has occurred
     */
    public static void close(Workbook workbook, CSVWriter writer) throws IOException {
        workbook.close();
        writer.close();
    }

    /**
     * Close a workbook
     * @param workbook The {@code Workbook} to close
     * @param fileOutputStream The {@code FileOutputStream} to close
     * @param reader The {@code CSVReader} to close
     * @throws IOException If an I/O error has occurred
     */
    public static void close(Workbook workbook, FileOutputStream fileOutputStream, CSVReader reader) throws IOException {
        workbook.close();
        fileOutputStream.close();
        reader.close();
    }
}
