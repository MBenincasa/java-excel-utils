package io.github.mbenincasa.javaexcelutils.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * This annotation is needed when the parseToObject() method in ExcelSheet is used
 * @author Mirko Benincasa
 * @since 0.5.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelCellMapping {

    /**
     * @return the value to add to the source row
     */
    int deltaRow();

    /**
     * @return the value to add to the source column
     */
    int deltaCol();
}
