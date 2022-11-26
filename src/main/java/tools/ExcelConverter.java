package tools;

import enums.ExcelExtension;

import java.io.File;
import java.io.IOException;
import java.util.List;

public interface ExcelConverter {

    File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz) throws IllegalAccessException, IOException;

    File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String filename) throws IllegalAccessException, IOException;

    File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String path, String filename) throws IllegalAccessException, IOException;

    File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String path, String filename, Boolean writeHeader) throws IllegalAccessException, IOException;

    File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String path, String filename, ExcelExtension extension) throws IllegalAccessException, IOException;

    File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String path, String filename, ExcelExtension extension, Boolean writeHeader) throws IllegalAccessException, IOException;
}
