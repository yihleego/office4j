package io.leego.office4j.excel.style;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author Yihleego
 */
public class MoneyStyle extends ValueStyle {

    @Override
    public CellStyle getCellStyle(Workbook workbook) {
        CellStyle cellStyle = super.getCellStyle(workbook);
        cellStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("0.00"));
        return cellStyle;
    }

}
