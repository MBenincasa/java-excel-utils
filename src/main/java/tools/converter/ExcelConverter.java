package tools.converter;

import java.io.File;
import java.io.IOException;
import java.util.List;

public interface ExcelConverter {

    File convertObjectsToExcelFile(List<? extends Object> objects, Class<? extends Object> clazz, String pathname, Boolean writeHeader) throws IllegalAccessException, IOException;
}
