package io.github.mbenincasa.javaexcelutils.model.converter;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

/**
 * This is a support class that is used by the {@code Converter} to perform conversions from Json to Excel
 * @author Mirko Benincasa
 * @since 0.4.0
 * @param <T> The class parameter, for each Sheet, that maps a Json object into the parameter object
 */
@AllArgsConstructor
@Getter
@Setter
public class JsonToExcel<T> {

    /**
     * The name to assign to the Sheet
     */
    private String sheetName;

    /**
     * The object class
     */
    private Class<T> clazz;
}
