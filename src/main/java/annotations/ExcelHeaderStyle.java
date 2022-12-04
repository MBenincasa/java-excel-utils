package annotations;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelHeaderStyle {

    IndexedColors cellColor() default IndexedColors.GREY_50_PERCENT;

    HorizontalAlignment horizontal() default HorizontalAlignment.LEFT;

    VerticalAlignment vertical() default VerticalAlignment.TOP;

    boolean autoSize() default false;
}
