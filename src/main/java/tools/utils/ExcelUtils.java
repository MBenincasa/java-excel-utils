package tools.utils;

import java.io.File;

public interface ExcelUtils {

    Integer countAllRows(File file, Boolean alsoEmptyRows) throws Exception;

    Integer countAllRows(File file, Boolean alsoEmptyRows, String sheetName) throws Exception;
}
