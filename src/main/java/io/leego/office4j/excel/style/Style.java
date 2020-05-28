package io.leego.office4j.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author Yihleego
 */
public interface Style {

    CellStyle getCellStyle(Workbook workbook);

}
