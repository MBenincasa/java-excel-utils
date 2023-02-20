package model;

import enums.Extension;
import exceptions.OpenWorkbookException;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;

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

    public void close() throws IOException {
        this.workbook.close();
    }
}
