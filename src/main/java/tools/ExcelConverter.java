package tools;

import enums.Extension;

import java.io.File;
import java.io.IOException;
import java.util.List;

public interface ExcelConverter {

    File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz) throws IllegalAccessException, IOException;

    File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String filename) throws IllegalAccessException, IOException;

    File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String path, String filename) throws IllegalAccessException, IOException;

    File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String path, String filename, Boolean writeHeader) throws IllegalAccessException, IOException;

    File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String path, String filename, Extension extension) throws IllegalAccessException, IOException;

    File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String path, String filename, Extension extension, Boolean writeHeader) throws IllegalAccessException, IOException;
}
