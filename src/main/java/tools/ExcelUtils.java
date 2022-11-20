package tools;

import exceptions.ExtensionNotValidException;

import java.io.File;
import java.io.IOException;

public interface ExcelUtils {

    Integer countAllRows(File file) throws ExtensionNotValidException, IOException;

    Integer countAllRows(File file, Boolean alsoEmptyRows) throws ExtensionNotValidException, IOException;

    Integer countAllRows(File file, Boolean alsoEmptyRows, String sheetName) throws ExtensionNotValidException, IOException;
}
