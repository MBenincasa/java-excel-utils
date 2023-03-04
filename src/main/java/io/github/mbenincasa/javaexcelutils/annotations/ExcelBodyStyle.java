package io.github.mbenincasa.javaexcelutils.annotations;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * This annotation defines the style of the Excel file body when converting a list of objects to an Excel file
 * @author Mirko Benincasa
 * @since 0.1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelBodyStyle {

    /**
     * @return the background color of the cell. The default is {@code IndexedColors.GREY_25_PERCENT}
     */
    IndexedColors cellColor() default IndexedColors.GREY_25_PERCENT;

    /**
     * @return the horizontal orientation of the text in the cell. The default is {@code HorizontalAlignment.LEFT}
     */
    HorizontalAlignment horizontal() default HorizontalAlignment.LEFT;

    /**
     * @return the vertical orientation of the text in the cell. The default is {@code VerticalAlignment.TOP}
     */
    VerticalAlignment vertical() default VerticalAlignment.TOP;
}
