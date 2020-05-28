package io.leego.office4j.excel.style;

import io.leego.office4j.util.DatePattern;

/**
 * @author Yihleego
 */
public class TimeStyle extends AbstractDateStyle {

    @Override
    public String getPattern() {
        return DatePattern.TIME;
    }

}
