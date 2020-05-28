package io.leego.office4j.excel;

import io.leego.office4j.excel.style.Style;
import org.apache.poi.ss.usermodel.CellStyle;

import java.lang.reflect.Field;

/**
 * @author Yihleego
 */
public class ExcelMetaField {
    private final Field field;
    private final String name;
    private int column;
    private int width;
    private final Style titleStyle;
    private final Style valueStyle;
    private final CellStyle titleCellStyle;
    private final CellStyle valueCellStyle;

    public ExcelMetaField(Field field, String name, int column, int width,
                          Style titleStyle, Style valueStyle, CellStyle titleCellStyle, CellStyle valueCellStyle) {
        this.field = field;
        this.name = name;
        this.column = column;
        this.width = width;
        this.titleStyle = titleStyle;
        this.valueStyle = valueStyle;
        this.titleCellStyle = titleCellStyle;
        this.valueCellStyle = valueCellStyle;
    }

    public Field getField() {
        return field;
    }

    public String getName() {
        return name;
    }

    public int getColumn() {
        return column;
    }

    public int getWidth() {
        return width;
    }

    public void setColumn(int column) {
        this.column = column;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public Style getTitleStyle() {
        return titleStyle;
    }

    public Style getValueStyle() {
        return valueStyle;
    }

    public CellStyle getTitleCellStyle() {
        return titleCellStyle;
    }

    public CellStyle getValueCellStyle() {
        return valueCellStyle;
    }
}
