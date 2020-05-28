# Getting Started

Please make sure that Java version is 1.8 and above.

# Usage

## Simple example

Create a sheet, and write data to OutputStream.

```java
try (ExcelBuilder builder = ExcelBuilder.builder()) {
    builder.createSheet("sheet1")
            .write(data)
            .toOutputStream(outputStream);
} catch (IOException e) {
    logger.error("", e);
}
```

## APIs

```java
try (ExcelBuilder builder = ExcelBuilder.builder()) {
    builder
            // Sets the number of rows that are kept in memory until flushed out.
            .rowAccessInMemory(1000)
            // Whether to use gzip compression for temporary files.
            .compressTmpFiles(true)
            // Whether to adjust the column width to fit the contents.
            .autoColumnWidth(true)
            // Creates a sheet with a name.
            .createSheet("foo")
            // Sets header.
            .header(sheet -> {
                SXSSFRow title = sheet.createRow(0);
                SXSSFCell sxssfCell = title.createCell(0);
                sxssfCell.setCellValue("header");
                return 1;
            })
            // Sets footer.
            .footer((sheet, rowNumber) -> {
                SXSSFRow title = sheet.createRow(rowNumber);
                SXSSFCell sxssfCell = title.createCell(0);
                sxssfCell.setCellValue("footer");
                return 1;
            })
            // Whether to display column title.
            .displayTitle(true)
            // Whether to keep title when partition sheets.
            .keepTitleWhenPartition(true)
            // Whether to keep header when partition sheets.
            .keepHeaderWhenPartition(true)
            // Whether to keep footer when partition sheets.
            .keepFooterWhenPartition(true)
            // Exports an excel without titles if data is empty and {@link ExcelSheetBuilder#beanClass} is null.
            // If titles are expected to exist, and {@link ExcelSheetBuilder#beanClass} is required.
            .beanClass(Foo.class)
            // Sets maximum row per sheet.
            .maxRowPerSheet(1000)
            // Writes data.
            .write(data)
            .and()
            .createSheet("bar")
            .write(writer -> {
                for (List<Bar> o : dataGroup) {
                    writer.append(o);
                }
            }).toFile("e:\\foobar.xlsx");
} catch (IOException e) {
    logger.error("", e);
}
```