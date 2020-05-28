package io.leego.office4j.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author Yihleego
 */
public final class NonStyle implements Style {

    @Override
    public CellStyle getCellStyle(Workbook workbook) {
        throw new UnsupportedOperationException("Don't touch me!");
    }

}
