package tools.interfaces;

import enums.ExcelExtension;
import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.IOException;

public interface ExcelWorkbookUtils {

    Workbook open(FileInputStream fileInputStream, String extension) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    Workbook create();

    Workbook create(String extension) throws ExtensionNotValidException;

    Workbook create(ExcelExtension extension);
}
