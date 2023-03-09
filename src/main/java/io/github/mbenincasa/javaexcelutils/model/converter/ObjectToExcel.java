package io.github.mbenincasa.javaexcelutils.model.converter;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

import java.util.stream.Stream;

/**
 * This is a helper class used by {@code Converter} to convert objects in an Excel Sheet
 * @author Mirko Benincasa
 * @since 0.4.0
 * @param <T> The class parameter, for each sheet, which maps objects to a Sheet
 */
@AllArgsConstructor
@Getter
@Setter
public class ObjectToExcel<T> {

    /**
     * The name to assign to the Sheet
     */
    private String sheetName;

    /**
     * The object class
     */
    private Class<T> clazz;

    /**
     * A Stream of objects
     */
    private Stream<T> stream;
}
