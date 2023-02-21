package model;

import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import enums.Extension;
import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import tools.ExcelUtility;

import java.io.*;

@AllArgsConstructor
@Getter
public class ExcelWorkbook {

    private Workbook workbook;

    public ExcelWorkbook(Extension extension) {
        switch (extension) {
            case XLS -> this.workbook = new HSSFWorkbook();
            case XLSX -> this.workbook = new XSSFWorkbook();
        }
    }

    public ExcelWorkbook(InputStream inputStream) throws OpenWorkbookException {
        try {
            this.workbook = new XSSFWorkbook(inputStream);
        } catch (OfficeXmlFileException | OLE2NotOfficeXmlFileException | IOException e) {
            try {
                this.workbook = new HSSFWorkbook(inputStream);
            } catch (IOException ex) {
                throw new OpenWorkbookException("The workbook could not be opened", ex);
            }
        }
    }

    public static ExcelWorkbook open(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Check extension */
        String extension = ExcelUtility.checkExcelExtension(file.getName());

        /* Open file input stream */
        FileInputStream fileInputStream = new FileInputStream(file);
        ExcelWorkbook excelWorkbook = open(fileInputStream, extension);

        /* Close the stream before return */
        fileInputStream.close();
        return excelWorkbook;
    }

    public static ExcelWorkbook open(InputStream inputStream, String extension) throws ExtensionNotValidException, IOException, OpenWorkbookException {
        /* Check the extension */
        if (!ExcelUtility.isValidExcelExtension(extension)) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }

        return new ExcelWorkbook(inputStream);
    }

    public static ExcelWorkbook create() {
        return create(Extension.XLSX);
    }

    public static ExcelWorkbook create(String extension) throws ExtensionNotValidException {
        if (!ExcelUtility.isValidExcelExtension(extension)) {
            throw new ExtensionNotValidException("Pass a file with the XLS or XLSX extension");
        }
        return create(Extension.getExcelExtension(extension));
    }

    public static ExcelWorkbook create(Extension extension) {
        return new ExcelWorkbook(extension);
    }

    public void close() throws IOException {
        this.workbook.close();
    }

    public void close(InputStream inputStream) throws IOException {
        this.workbook.close();
        inputStream.close();
    }

    public void close(OutputStream outputStream) throws IOException {
        this.workbook.close();
        outputStream.close();
    }

    public void close(OutputStream outputStream, InputStream inputStream) throws IOException {
        this.workbook.close();
        inputStream.close();
        outputStream.close();
    }

    public void close(CSVWriter writer) throws IOException {
        this.workbook.close();
        writer.close();
    }

    public void close(OutputStream outputStream, CSVReader reader) throws IOException {
        this.workbook.close();
        outputStream.close();
        reader.close();
    }
}
