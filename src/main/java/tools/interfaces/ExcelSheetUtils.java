package tools.interfaces;

import exceptions.ExtensionNotValidException;
import exceptions.OpenWorkbookException;
import exceptions.SheetNotFoundException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.util.List;

public interface ExcelSheetUtils {

    Integer countAll(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    List<String> getAllNames(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    Integer getIndex(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    String getNameByIndex(File file, Integer position) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    Sheet create(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    Sheet create(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException;

    Sheet create(Workbook workbook);

    Sheet create(Workbook workbook, String sheetName);

    Sheet open(File file) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    Sheet open(File file, String sheetName) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    Sheet open(File file, Integer position) throws ExtensionNotValidException, IOException, OpenWorkbookException, SheetNotFoundException;

    Sheet open(Workbook workbook) throws SheetNotFoundException;

    Sheet open(Workbook workbook, String sheetName) throws SheetNotFoundException;

    Sheet open(Workbook workbook, Integer position) throws SheetNotFoundException;

    Sheet openOrCreate(Workbook workbook, String sheetName);
}
