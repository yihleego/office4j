package io.leego.office4j.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;

/**
 * @author Yihleego
 */
public interface ExcelFooter {

    /**
     * Returns the height occupied by the footer.
     * @param sheet     {@link SXSSFSheet} sheet
     * @param rowNumber the number of current rows.
     * @return the height occupied by the footer.
     */
    int apply(SXSSFSheet sheet, int rowNumber);

}
