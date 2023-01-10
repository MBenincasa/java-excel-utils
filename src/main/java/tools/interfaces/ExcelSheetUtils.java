package tools.interfaces;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;

import java.io.File;
import java.io.IOException;
import java.util.List;

public interface ExcelSheetUtils {

    Integer countAll(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    List<String> getAllNames(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    Integer getIndex(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    String getNameByIndex(File file, Integer position) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;
}
