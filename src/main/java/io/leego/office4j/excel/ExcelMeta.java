package io.leego.office4j.excel;

import io.leego.office4j.excel.style.NonStyle;
import io.leego.office4j.excel.style.Style;

import java.lang.annotation.*;

/**
 * @author Yihleego
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.METHOD})
@Documented
public @interface ExcelMeta {

    /**
     * Title name.
     * Alias for {@link ExcelMeta#name}.
     */
    String value() default "";

    /**
     * Title name.
     * Alias for {@link ExcelMeta#value}.
     */
    String name() default "";

    /** Zero-based column index. */
    int column() default -1;

    /** Column width, the default value equals column title's length. */
    int width() default -1;

    /** Title style. */
    Class<? extends Style> titleStyle() default NonStyle.class;

    /** Value style. */
    Class<? extends Style> valueStyle() default NonStyle.class;
}
