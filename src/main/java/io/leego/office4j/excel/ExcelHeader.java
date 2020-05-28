package io.leego.office4j.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;

/**
 * @author Yihleego
 */
public interface ExcelHeader {

    /**
     * Returns the height occupied by the header.
     * @param sheet {@link SXSSFSheet} sheet
     * @return the height occupied by the header.
     */
    int apply(SXSSFSheet sheet);

}
