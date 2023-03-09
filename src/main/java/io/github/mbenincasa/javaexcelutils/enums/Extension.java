package io.github.mbenincasa.javaexcelutils.enums;

import io.github.mbenincasa.javaexcelutils.exceptions.ExtensionNotValidException;
import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.Arrays;
import java.util.Optional;

/**
 * This Enum defines the file extensions supported by the library.
 * @author Mirko Benincasa
 * @since 0.1.0
 */
@Getter
@AllArgsConstructor
public enum Extension {

    /**
     * This extension is used for a Microsoft Office spreadsheet up to version 2003
     */
    XLS("xls", "EXCEL"),

    /**
     * This extension is used for a Microsoft Office spreadsheet from version 2007 onwards
     */
    XLSX("xlsx", "EXCEL"),

    /**
     * This extension is used for CSV (Comma-separated values) files
     */
    CSV("csv", "CSV"),

    /**
     * @since 0.4.0
     * This extension is used for JSON (JavaScript Object Notation) files
     */
    JSON("json", "JSON");

    /**
     * The extension's name
     */
    private final String ext;

    /**
     * The extension's type
     */
    private final String type;

    /**
     * This method retrieves the Enum value based on the extension name provided as input
     * @param ext the name of an Excel file extension to search for
     * @return The Enum value that matches the name provided as input and is of type Excel
     * @throws ExtensionNotValidException if no Enum value is found
     */
    public static Extension getExcelExtension(String ext) throws ExtensionNotValidException {
        Optional<Extension> extensionOptional = Arrays.stream(Extension.values()).filter(e -> ext.equalsIgnoreCase(e.getExt()) && e.getType().equals("EXCEL")).findFirst();
        if (extensionOptional.isEmpty()) {
            throw new ExtensionNotValidException();
        }
        return extensionOptional.get();
    }

    /**
     * @return {@code true} if it has type EXCEL
     * @since 0.1.1
     */
    public Boolean isExcelExtension() {
        return this.type.equalsIgnoreCase("EXCEL");
    }
}
