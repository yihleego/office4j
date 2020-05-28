package io.leego.office4j.excel;

import io.leego.office4j.excel.exception.ExcelException;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Yihleego
 */
public class ExcelBuilder implements AutoCloseable {
    private static final int ROW_ACCESS_WINDOW_SIZE = 1 << 10;
    private SXSSFWorkbook workbook;
    private List<ExcelSheetBuilder> sheets;
    private int rowAccessInMemory;
    private boolean compressTmpFiles = true;
    private boolean autoColumnWidth = false;

    public ExcelBuilder() {
        this(ROW_ACCESS_WINDOW_SIZE);
    }

    public ExcelBuilder(int rowAccessInMemory) {
        this.rowAccessInMemory = rowAccessInMemory;
    }

    public static ExcelBuilder builder() {
        return new ExcelBuilder();
    }

    /**
     * Sets the number of rows that are kept in memory until flushed out.
     * @param rowAccessInMemory the number of rows that are kept in memory until flushed out.
     */
    public ExcelBuilder rowAccessInMemory(int rowAccessInMemory) {
        this.rowAccessInMemory = rowAccessInMemory;
        return this;
    }

    /**
     * Whether to use gzip compression for temporary files.
     * @param compressTmpFiles whether to compress temp files
     */
    public ExcelBuilder compressTmpFiles(boolean compressTmpFiles) {
        this.compressTmpFiles = compressTmpFiles;
        if (this.workbook != null) {
            this.workbook.setCompressTempFiles(compressTmpFiles);
        }
        return this;
    }

    /**
     * Whether to adjust the column width to fit the contents.
     * @param autoColumnWidth whether to adjust width.
     */
    public ExcelBuilder autoColumnWidth(boolean autoColumnWidth) {
        this.autoColumnWidth = autoColumnWidth;
        return this;
    }

    /**
     * Creates a sheet without a name.
     */
    public ExcelSheetBuilder createSheet() {
        return createSheet(null);
    }

    /**
     * Creates a sheet with a name.
     * @param name the sheet name.
     */
    public ExcelSheetBuilder createSheet(String name) {
        ensureInitialization();
        ExcelSheetBuilder sheetBuilder = new ExcelSheetBuilder(this, this.workbook, name, this.sheets.size());
        this.sheets.add(sheetBuilder);
        return sheetBuilder;
    }

    /**
     * Writes data to {@link OutputStream}.
     * @param outputStream {@link OutputStream}.
     */
    public void toOutputStream(OutputStream outputStream) throws IOException {
        if (outputStream == null) {
            throw new NullPointerException("OutputStream is required.");
        }
        finish(outputStream, true);
    }

    /**
     * Writes data to {@link OutputStream}.
     * @param outputStream      {@link OutputStream}.
     * @param closeOutputStream whether to OutputStream.
     */
    public void toOutputStream(OutputStream outputStream, boolean closeOutputStream) throws IOException {
        if (outputStream == null) {
            throw new NullPointerException("OutputStream is required.");
        }
        finish(outputStream, closeOutputStream);
    }

    /**
     * Writes data to file.
     * @param path the file path.
     */
    public void toFile(String path) throws IOException {
        if (path == null) {
            throw new NullPointerException("Path is required.");
        }
        toFile(new File(path));
    }

    /**
     * Writes data to file.
     * @param file the file.
     */
    public void toFile(File file) throws IOException {
        if (file == null) {
            throw new NullPointerException("File is required.");
        }
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        finish(fileOutputStream, true);
    }

    private void finish(OutputStream outputStream, boolean close) throws IOException {
        if (this.workbook == null) {
            throw new ExcelException("Nothing to output");
        }
        // Write footer
        this.sheets.forEach(ExcelSheetBuilder::writeFooter);
        // Write excel to OutputStream
        this.workbook.write(outputStream);
        outputStream.flush();
        if (close) {
            outputStream.close();
        }
        this.workbook.close();
    }

    @Override
    public void close() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
        if (sheets != null) {
            sheets.clear();
        }
    }

    private void ensureInitialization() {
        if (this.workbook == null) {
            this.workbook = new SXSSFWorkbook(null, rowAccessInMemory, compressTmpFiles);
        }
        if (this.sheets == null) {
            this.sheets = new ArrayList<>();
        }
    }

    protected SXSSFWorkbook getWorkbook() {
        return this.workbook;
    }

    public boolean isAutoColumnWidth() {
        return autoColumnWidth;
    }

}
