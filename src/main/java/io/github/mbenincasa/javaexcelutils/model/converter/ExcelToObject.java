package io.github.mbenincasa.javaexcelutils.model.converter;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

/**
 * This is a support class that is used by the {@code Converter} to perform conversions from Excel to Objects
 * @author Mirko Benincasa
 * @since 0.4.0
 * @param <T> The class parameter, for each Sheet, that maps a row into the parameter object
 */
@AllArgsConstructor
@Getter
@Setter
public class ExcelToObject<T> {

    /**
     * The name of the Sheet to read
     */
    private String sheetName;

    /**
     * The object class
     */
    private Class<T> clazz;
}
