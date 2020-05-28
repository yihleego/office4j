package io.leego.office4j.excel;

import io.leego.office4j.excel.exception.ExcelException;
import io.leego.office4j.excel.style.NonStyle;
import io.leego.office4j.excel.style.Style;
import io.leego.office4j.util.Adder;
import io.leego.office4j.util.ReflectUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.List;
import java.util.function.Consumer;

/**
 * @author Yihleego
 */
public class ExcelSheetBuilder {
    private final SXSSFWorkbook workbook;
    private final ExcelBuilder parent;
    private ExcelHeader header;
    private ExcelFooter footer;
    private List<ExcelMetaField> metaFields;
    private List<ExcelSheet> subSheets;
    private ExcelSheet curSheet;
    private final String name;
    private final int index;
    private int subIndex;
    private int maxRowPerSheet;
    private boolean displayTitle = true;
    private boolean keepTitleWhenPartition = true;
    private boolean keepHeaderWhenPartition = true;
    private boolean keepFooterWhenPartition = true;
    private Class<?> beanClass;

    ExcelSheetBuilder(ExcelBuilder parent, SXSSFWorkbook workbook, String name, int index) {
        this.parent = parent;
        this.name = name;
        this.index = index;
        this.workbook = workbook;
        this.curSheet = new ExcelSheet(workbook.createSheet(buildName()), true);
        this.subSheets = new ArrayList<>();
        this.subSheets.add(this.curSheet);
        this.subIndex = 0;
    }

    /**
     * Returns {@link ExcelBuilder}
     * @return {@link ExcelBuilder}
     */
    public ExcelBuilder and() {
        return this.parent;
    }

    /**
     * Sets maximum row per sheet.
     * @param maxRowPerSheet the maximum row per sheet.
     */
    public ExcelSheetBuilder maxRowPerSheet(int maxRowPerSheet) {
        this.maxRowPerSheet = maxRowPerSheet;
        return this;
    }

    /**
     * Whether to display column title.
     * @param displayTitle wWhether to display column title.
     */
    public ExcelSheetBuilder displayTitle(boolean displayTitle) {
        this.displayTitle = displayTitle;
        return this;
    }

    /**
     * Whether to keep title when partition sheets.
     * @param keepTitleWhenPartition wWhether to keep title when partition sheets.
     */
    public ExcelSheetBuilder keepTitleWhenPartition(boolean keepTitleWhenPartition) {
        this.keepTitleWhenPartition = keepTitleWhenPartition;
        return this;
    }

    /**
     * Whether to keep header when partition sheets.
     * @param keepHeaderWhenPartition wWhether to keep header when partition sheets.
     */
    public ExcelSheetBuilder keepHeaderWhenPartition(boolean keepHeaderWhenPartition) {
        this.keepHeaderWhenPartition = keepHeaderWhenPartition;
        return this;
    }

    /**
     * Whether to keep footer when partition sheets.
     * @param keepFooterWhenPartition wWhether to keep footer when partition sheets.
     */
    public ExcelSheetBuilder keepFooterWhenPartition(boolean keepFooterWhenPartition) {
        this.keepFooterWhenPartition = keepFooterWhenPartition;
        return this;
    }

    /**
     * Sets header.
     * @param header {@link ExcelHeader}
     */
    public ExcelSheetBuilder header(ExcelHeader header) {
        if (header == null) {
            return this;
        }
        this.header = header;
        return this;
    }

    /**
     * Sets footer.
     * @param footer {@link ExcelFooter}
     */
    public ExcelSheetBuilder footer(ExcelFooter footer) {
        this.footer = footer;
        return this;
    }

    /**
     * Exports an excel without titles if data is empty and {@link ExcelSheetBuilder#beanClass} is null.
     * If titles are expected to exist, and {@link ExcelSheetBuilder#beanClass} is required.
     * @param beanClass the bean class.
     */
    public ExcelSheetBuilder beanClass(Class<?> beanClass) {
        this.beanClass = beanClass;
        return this;
    }

    /**
     * Writes data.
     * @param data the data.
     */
    public ExcelSheetBuilder write(Collection<?> data) {
        try {
            writeData(data);
        } catch (Exception e) {
            throw new ExcelException(e);
        }
        return this;
    }

    /**
     * Writes data.
     * @param action the action.
     */
    public ExcelSheetBuilder write(Consumer<Writer> action) {
        if (action == null) {
            return this;
        }
        action.accept(new Writer(this));
        return this;
    }

    protected void writeData(Collection<?> data) throws InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        writeHeader(curSheet);
        writeTitle(data, curSheet);
        if (data == null || data.isEmpty()) {
            return;
        }
        if (maxRowPerSheet > 0 && curSheet.getDataSize() + data.size() > maxRowPerSheet) {
            int currentRest = maxRowPerSheet - curSheet.getDataSize();
            if (currentRest != 0) {
                writeData(data, 0, currentRest, curSheet);
            }
            int nextRest = data.size() - currentRest;
            if (nextRest > 0) {
                int requiredSheetNumber = nextRest % maxRowPerSheet == 0
                        ? nextRest / maxRowPerSheet
                        : nextRest / maxRowPerSheet + 1;
                for (int i = 0; i < requiredSheetNumber; i++) {
                    String sheetName = buildName() + (++subIndex);
                    curSheet = new ExcelSheet(workbook.createSheet(sheetName));
                    subSheets.add(curSheet);
                    if (keepHeaderWhenPartition) {
                        writeHeader(curSheet);
                    }
                    if (keepTitleWhenPartition) {
                        writeTitle(data, curSheet);
                    }
                    writeData(data, i * maxRowPerSheet + currentRest, maxRowPerSheet, curSheet);
                }
            }
        } else {
            writeData(data, 0, data.size(), curSheet);
        }
    }

    protected void writeData(Collection<?> data, int offset, int rows, ExcelSheet sheet) {
        Adder index = new Adder();
        data.stream().skip(offset).limit(rows).forEach(o -> {
            SXSSFRow row = sheet.getSheet().createRow(sheet.getCurrentRow() + index.getAndIncrement());
            try {
                for (ExcelMetaField metaField : metaFields) {
                    SXSSFCell cell = row.createCell(metaField.getColumn());
                    // Set style
                    if (metaField.getValueCellStyle() != null) {
                        cell.setCellStyle(metaField.getValueCellStyle());
                    }
                    // Set value
                    Object value = ReflectUtils.getFieldValue(o, metaField.getField(), true);
                    if (value == null) {
                        cell.setCellType(CellType.BLANK);
                    } else if (value instanceof Double) {
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue((double) value);
                    } else if (value instanceof Integer) {
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue((int) value);
                    } else if (value instanceof Long) {
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue((long) value);
                    } else if (value instanceof Float) {
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue((float) value);
                    } else if (value instanceof BigDecimal) {
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue(((BigDecimal) value).doubleValue());
                    } else if (value instanceof LocalDateTime) {
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue((LocalDateTime) value);
                    } else if (value instanceof LocalDate) {
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue((LocalDate) value);
                    }/* else if (value instanceof LocalTime) {
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue(LocalDateTime.of(DATE, (LocalTime) value));
                    }*/ else if (value instanceof Date) {
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue((Date) value);
                    } else if (value instanceof Calendar) {
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue((Calendar) value);
                    } else if (value instanceof RichTextString) {
                        cell.setCellType(CellType.STRING);
                        cell.setCellValue((RichTextString) value);
                    } else {
                        cell.setCellType(CellType.STRING);
                        cell.setCellValue(value.toString());
                    }
                }
            } catch (IllegalAccessException e) {
                throw new ExcelException(e);
            }
        });
        sheet.setDataSize(sheet.getDataSize() + index.get());
        sheet.setCurrentRow(sheet.getCurrentRow() + index.get());
    }

    protected void writeTitle(Collection<?> data, ExcelSheet sheet) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException, InstantiationException {
        if (sheet.isWrittenTitle()) {
            return;
        }
        if (beanClass == null && (data == null || data.isEmpty())) {
            return;
        }
        Class<?> targetType = null;
        if (beanClass != null) {
            targetType = beanClass;
        } else {
            for (Object o : data) {
                if (o != null) {
                    targetType = o.getClass();
                    break;
                }
            }
        }
        if (targetType == null) {
            return;
        }
        // Global styles
        Style globalTitleStyle = null;
        Style globalValueStyle = null;
        CellStyle globalTitleCellStyle = null;
        CellStyle globalValueCellStyle = null;
        ExcelStyle excelStyle = ReflectUtils.getAnnotation(targetType, ExcelStyle.class);
        if (excelStyle != null) {
            if (excelStyle.titleStyle() != NonStyle.class) {
                globalTitleStyle = excelStyle.titleStyle().getConstructor().newInstance();
                globalTitleCellStyle = globalTitleStyle.getCellStyle(workbook);
            }
            if (excelStyle.valueStyle() != NonStyle.class) {
                globalValueStyle = excelStyle.valueStyle().getConstructor().newInstance();
                globalValueCellStyle = globalValueStyle.getCellStyle(workbook);
            }
        }

        boolean naturalOrder = true;
        int maxOrder = 0;
        List<ExcelMetaField> metaFields = new ArrayList<>();
        Field[] fields = ReflectUtils.getFields(targetType);
        for (Field field : fields) {
            ExcelMeta excelMeta = ReflectUtils.getAnnotation(field, ExcelMeta.class);
            if (excelMeta == null) {
                continue;
            }
            if (excelMeta.column() > maxOrder) {
                maxOrder = excelMeta.column();
            }
            if (naturalOrder && excelMeta.column() != -1) {
                naturalOrder = false;
            }
            String name = excelMeta.value().length() > 0 ? excelMeta.value() : excelMeta.name();
            Style titleStyle;
            Style valueStyle;
            CellStyle titleCellStyle;
            CellStyle valueCellStyle;
            excelMeta.value();
            if (excelMeta.titleStyle() != NonStyle.class) {
                titleStyle = excelMeta.titleStyle().getConstructor().newInstance();
                titleCellStyle = titleStyle.getCellStyle(workbook);
            } else {
                titleStyle = globalTitleStyle;
                titleCellStyle = globalTitleCellStyle;
            }
            if (excelMeta.valueStyle() != NonStyle.class) {
                valueStyle = excelMeta.valueStyle().getConstructor().newInstance();
                valueCellStyle = valueStyle.getCellStyle(workbook);
            } else {
                valueStyle = globalValueStyle;
                valueCellStyle = globalValueCellStyle;
            }
            metaFields.add(new ExcelMetaField(
                    field,
                    name,
                    excelMeta.column(),
                    excelMeta.width(),
                    titleStyle,
                    valueStyle,
                    titleCellStyle,
                    valueCellStyle
            ));
        }
        if (metaFields.isEmpty()) {
            return;
        }
        // Sort
        if (naturalOrder) {
            for (int i = 0; i < metaFields.size(); i++) {
                metaFields.get(i).setColumn(i);
            }
        } else {
            metaFields.sort((o1, o2) -> {
                if (o1.getColumn() == o2.getColumn()) {
                    return 0;
                } else if (o1.getColumn() == -1) {
                    return 1;
                } else if (o2.getColumn() == -1) {
                    return -1;
                } else {
                    return o1.getColumn() - o2.getColumn();
                }
            });
            for (ExcelMetaField metaField : metaFields) {
                if (metaField.getColumn() < 0) {
                    metaField.setColumn(++maxOrder);
                }
            }
        }
        if (this.displayTitle) {
            int rowIndex = Math.max(0, sheet.getCurrentRow());
            SXSSFRow row = sheet.getSheet().createRow(rowIndex);
            for (ExcelMetaField metaField : metaFields) {
                SXSSFCell cell = row.createCell(metaField.getColumn(), CellType.STRING);
                // Set style
                if (metaField.getTitleCellStyle() != null) {
                    cell.setCellStyle(metaField.getTitleCellStyle());
                }
                // Set value
                cell.setCellValue(metaField.getName());
                // Obtain column width
                int width = metaField.getWidth() > 0
                        ? metaField.getWidth() * 256
                        : (metaField.getName().length() * 2 + 4) * 256;
                sheet.getSheet().setColumnWidth(metaField.getColumn(), width);
            }
            sheet.setCurrentRow(rowIndex + 1);
            sheet.setWrittenTitle(true);
        }
        this.metaFields = metaFields;
    }

    protected void writeHeader(ExcelSheet sheet) {
        if (header != null && !sheet.isWrittenHeader()) {
            int headerHeight = header.apply(sheet.getSheet());
            sheet.setCurrentRow(sheet.getCurrentRow() + headerHeight);
            sheet.setWrittenHeader(true);
        }
    }

    protected void writeFooter() {
        for (ExcelSheet sheet : subSheets) {
            if (footer != null && !sheet.isWrittenFooter()
                    && (sheet.isPrimary() || keepFooterWhenPartition)) {
                int footerHeight = footer.apply(sheet.getSheet(), sheet.getCurrentRow());
                sheet.setCurrentRow(sheet.getCurrentRow() + footerHeight);
                sheet.setWrittenFooter(true);
            }
            if (parent.isAutoColumnWidth()) {
                for (ExcelMetaField metaField : metaFields) {
                    sheet.getSheet().trackAllColumnsForAutoSizing();
                    sheet.getSheet().autoSizeColumn(metaField.getColumn());
                }
            }
        }
    }

    /**
     * Writes data to {@link OutputStream}.
     * @param outputStream {@link OutputStream}.
     */
    public void toOutputStream(OutputStream outputStream) throws IOException {
        parent.toOutputStream(outputStream);
    }

    /**
     * Writes data to {@link OutputStream}.
     * @param outputStream      {@link OutputStream}.
     * @param closeOutputStream whether to OutputStream.
     */
    public void toOutputStream(OutputStream outputStream, boolean closeOutputStream) throws IOException {
        parent.toOutputStream(outputStream, closeOutputStream);
    }

    /**
     * Writes data to file.
     * @param path the file path.
     */
    public void toFile(String path) throws IOException {
        parent.toFile(path);
    }

    /**
     * Writes data to file.
     * @param file the file.
     */
    public void toFile(File file) throws IOException {
        parent.toFile(file);
    }

    private String buildName() {
        return name != null ? name : "Sheet" + index;
    }

    public static class Writer {
        private final ExcelSheetBuilder sheetBuilder;

        private Writer(ExcelSheetBuilder sheetBuilder) {
            this.sheetBuilder = sheetBuilder;
        }

        public Writer append(Collection<?> data) {
            sheetBuilder.write(data);
            return this;
        }
    }

}
