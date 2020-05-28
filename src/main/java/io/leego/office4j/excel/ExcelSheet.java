package io.leego.office4j.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;

/**
 * @author Yihleego
 */
public class ExcelSheet {
    private final SXSSFSheet sheet;
    private int currentRow;
    private int dataSize;
    private boolean writtenTitle;
    private boolean writtenHeader;
    private boolean writtenFooter;
    private final boolean primary;

    public ExcelSheet(SXSSFSheet sheet) {
        this.sheet = sheet;
        this.currentRow = 0;
        this.dataSize = 0;
        this.writtenTitle = false;
        this.writtenHeader = false;
        this.writtenFooter = false;
        this.primary = false;
    }

    public ExcelSheet(SXSSFSheet sheet, boolean primary) {
        this.sheet = sheet;
        this.currentRow = 0;
        this.dataSize = 0;
        this.writtenTitle = false;
        this.writtenHeader = false;
        this.writtenFooter = false;
        this.primary = primary;
    }

    public SXSSFSheet getSheet() {
        return sheet;
    }

    public int getCurrentRow() {
        return currentRow;
    }

    public void setCurrentRow(int currentRow) {
        this.currentRow = currentRow;
    }

    public int getDataSize() {
        return dataSize;
    }

    public void setDataSize(int dataSize) {
        this.dataSize = dataSize;
    }

    public boolean isWrittenTitle() {
        return writtenTitle;
    }

    public void setWrittenTitle(boolean writtenTitle) {
        this.writtenTitle = writtenTitle;
    }

    public boolean isWrittenHeader() {
        return writtenHeader;
    }

    public void setWrittenHeader(boolean writtenHeader) {
        this.writtenHeader = writtenHeader;
    }

    public boolean isWrittenFooter() {
        return writtenFooter;
    }

    public void setWrittenFooter(boolean writtenFooter) {
        this.writtenFooter = writtenFooter;
    }

    public boolean isPrimary() {
        return primary;
    }

}
