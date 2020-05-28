package io.leego.office4j.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author Yihleego
 */
public abstract class AbstractDateStyle extends ValueStyle {

    public abstract String getPattern();

    @Override
    public CellStyle getCellStyle(Workbook workbook) {
        CellStyle cellStyle = super.getCellStyle(workbook);
        DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat(getPattern()));
        return cellStyle;
    }

}
