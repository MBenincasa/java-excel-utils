package tools;

import enums.ExcelExtension;
import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public interface ExcelUtils {

    Integer countAllRows(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    Integer countAllRows(File file, Boolean alsoEmptyRows) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    Integer countAllRows(File file, Boolean alsoEmptyRows, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    Workbook openWorkbook(FileInputStream fileInputStream, String extension) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    Workbook createWorkbook();

    Workbook createWorkbook(String extension) throws ExtensionNotValidException;

    Workbook createWorkbook(ExcelExtension extension);

    Boolean isValidExcelExtension(String extension);
}
