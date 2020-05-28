package io.leego.office4j.excel.style;

import io.leego.office4j.util.DatePattern;

/**
 * @author Yihleego
 */
public class DateStyle extends AbstractDateStyle {

    @Override
    public String getPattern() {
        return DatePattern.DATE;
    }

}
