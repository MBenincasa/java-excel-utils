package enums;

import exceptions.ExtensionNotValidException;
import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.Arrays;
import java.util.Optional;

@Getter
@AllArgsConstructor
public enum ExcelExtension {

    XLS("xls"),
    XLSX("xlsx");

    private final String ext;

    public static ExcelExtension getExcelExtension(String ext) throws ExtensionNotValidException {
        Optional<ExcelExtension> extensionOptional = Arrays.stream(ExcelExtension.values()).filter(e -> ext.equalsIgnoreCase(e.getExt())).findFirst();
        if (extensionOptional.isEmpty()) {
            throw new ExtensionNotValidException();
        }
        return extensionOptional.get();
    }
}
