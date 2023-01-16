package enums;

import exceptions.ExtensionNotValidException;
import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.Arrays;
import java.util.Optional;

@Getter
@AllArgsConstructor
public enum Extension {

    XLS("xls", "EXCEL"),
    XLSX("xlsx", "EXCEL"),
    CSV("csv", "CSV");

    private final String ext;
    private final String type;

    public static Extension getExcelExtension(String ext) throws ExtensionNotValidException {
        Optional<Extension> extensionOptional = Arrays.stream(Extension.values()).filter(e -> ext.equalsIgnoreCase(e.getExt()) && e.getType().equals("EXCEL")).findFirst();
        if (extensionOptional.isEmpty()) {
            throw new ExtensionNotValidException();
        }
        return extensionOptional.get();
    }
}
