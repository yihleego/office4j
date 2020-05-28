package io.leego.office4j.excel;

import io.leego.office4j.excel.style.NonStyle;
import io.leego.office4j.excel.style.Style;

import java.lang.annotation.*;

/**
 * @author Yihleego
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
@Documented
public @interface ExcelStyle {

    /** Global title style. */
    Class<? extends Style> titleStyle() default NonStyle.class;

    /** Global value style. */
    Class<? extends Style> valueStyle() default NonStyle.class;

}
