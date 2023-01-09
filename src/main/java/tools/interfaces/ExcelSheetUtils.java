package tools.interfaces;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;

import java.io.File;
import java.io.IOException;
import java.util.List;

public interface ExcelSheetUtils {

    Integer countAllSheets(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    List<String> getAllSheetNames(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    Integer getSheetIndex(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    String getSheetNameAtPosition(File file, Integer position) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;
}
