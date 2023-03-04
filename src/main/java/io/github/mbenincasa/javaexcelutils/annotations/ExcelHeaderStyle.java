package io.github.mbenincasa.javaexcelutils.annotations;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * This annotation defines the header style of the Excel file when converting a list of objects to an Excel file
 * @author Mirko Benincasa
 * @since 0.1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelHeaderStyle {

    /**
     * @return the background color of the cell. The default is {@code IndexedColors.GREY_50_PERCENT}
     */
    IndexedColors cellColor() default IndexedColors.GREY_50_PERCENT;

    /**
     * @return the horizontal orientation of the text in the cell. The default is {@code HorizontalAlignment.LEFT}
     */
    HorizontalAlignment horizontal() default HorizontalAlignment.LEFT;

    /**
     * @return the vertical orientation of the text in the cell. The default is {@code VerticalAlignment.TOP}
     */
    VerticalAlignment vertical() default VerticalAlignment.TOP;

    /**
     * @return {@code true} if the autosize rules should be applied, otherwise {@code false}. The default is {@code false}
     */
    boolean autoSize() default false;
}
