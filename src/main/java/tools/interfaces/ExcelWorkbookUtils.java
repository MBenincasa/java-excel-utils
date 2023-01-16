package tools.interfaces;

import com.opencsv.CSVWriter;
import enums.Extension;
import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

public interface ExcelWorkbookUtils {

    Workbook open(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    Workbook open(FileInputStream fileInputStream, String extension) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    Workbook create();

    Workbook create(String extension) throws ExtensionNotValidException;

    Workbook create(Extension extension);

    void close(Workbook workbook) throws IOException;

    void close(Workbook workbook, FileInputStream fileInputStream) throws IOException;

    void close(Workbook workbook, FileOutputStream fileOutputStream) throws IOException;

    void close(Workbook workbook, FileOutputStream fileOutputStream, FileInputStream fileInputStream) throws IOException;

    void close(Workbook workbook, FileInputStream fileInputStream, CSVWriter writer) throws IOException;
}
