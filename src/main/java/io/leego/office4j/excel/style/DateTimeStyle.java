package io.leego.office4j.excel.style;

import io.leego.office4j.util.DatePattern;

/**
 * @author Yihleego
 */
public class DateTimeStyle extends AbstractDateStyle {

    @Override
    public String getPattern() {
        return DatePattern.DATE_TIME;
    }

}
