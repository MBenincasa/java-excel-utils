package tools.interfaces;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;

import java.io.File;
import java.io.IOException;

public interface ExcelUtils {

    Integer countAllRows(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    Integer countAllRows(File file, Boolean alsoEmptyRows) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    Integer countAllRows(File file, Boolean alsoEmptyRows, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    Boolean isValidExcelExtension(String extension);
}
