package naganaga.ss.writer;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import naganaga.ss.annotations.BodyStyle;
import naganaga.ss.annotations.Column;
import naganaga.ss.annotations.HeaderStyle;
import naganaga.ss.spec.SpreadSheetRowSpec;
import naganaga.ss.spec.SpreadSheetSpec;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UncheckedIOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.IntStream;

@Slf4j
public class SpreadSheetWriter<T> implements AutoCloseable {

    private OutputStream out;
    private SXSSFWorkbook wb;
    private Sheet sheet;
    private SpreadSheetSpec spreadSheetSpec;
    private int currentRow;
    private int bodyIndex = 0;

    private SpreadSheetRowSpec<T> rowSpec;
    private StyleContext context;

    public SpreadSheetWriter(OutputStream out, Class<T> spec) {
        this.out = out;
        wb = new SXSSFWorkbook(new XSSFWorkbook());

        spreadSheetSpec = new SpreadSheetSpec(spec);
        String sheetName = spreadSheetSpec.getSheetName();
        sheet = wb.getSheet(sheetName) == null ? wb.createSheet(sheetName) : wb.getSheet(sheetName);
        currentRow = spreadSheetSpec.getStartRow();

        this.rowSpec = new SpreadSheetRowSpec<>(spec);

        setColWidth();
        createStyleContext();
        if (spreadSheetSpec.hasHeader()) {
            writeHeader();
        }

    }

    private void setColWidth() {
        int offset = spreadSheetSpec.getStartCol();
        List<SpreadSheetRowSpec.ColumnSpec> columnSpecs = this.rowSpec.getColumnSpecs();
        IntStream.range(0, columnSpecs.size()).forEach(i -> {
            SpreadSheetRowSpec.ColumnSpec cs = columnSpecs.get(i);
            if (cs.getWidth() != -1) {
                sheet.setColumnWidth(offset + i, cs.getWidth());
            }
        });
    }

    private void createStyleContext() {
        StyleContext context = new StyleContext();
        Optional.ofNullable(spreadSheetSpec.getHeaderStyle()).ifPresent(style -> {
            CellStyle base = createHeaderStyle(style);
            context.setHeader(base);
            context.setHeaderHeight(style.height());
            context.setHeaderSurroundStyle(style.borderStyle());
            context.setHeaderSurroundColor(style.borderColor());
        });
        Optional.ofNullable(spreadSheetSpec.getBodyStyle()).ifPresent(style -> {
            CellStyle base = createBodyStyle(style);
            context.setBody(base);
            context.setBodyHeight(style.height());
            context.setBodySurroundStyle(style.borderStyle());
            context.setBodySurroundColor(style.borderColor());
        });

        Map<Integer, CellStyle> overrideHeader = new HashMap<>();
        Map<Integer, CellStyle> overrideBody = new HashMap<>();
        List<SpreadSheetRowSpec.ColumnSpec> rowSpecs = this.rowSpec.getColumnSpecs();
        IntStream.range(0, rowSpecs.size()).forEach(i -> {
            SpreadSheetRowSpec.ColumnSpec cs = rowSpecs.get(i);
            Optional.ofNullable(cs.getHeaderStyle()).ifPresent(style -> {
                CellStyle cellStyle = Optional.ofNullable(context.getHeader()).map(cStyle -> {
                    CellStyle override = createCellStyle();
                    override.cloneStyleFrom(cStyle);
                    return override;
                }).orElse(createCellStyle());
                cellStyle.setWrapText(style.isWrap());
                cellStyle.setFillForegroundColor(style.backgroundColor().getIndex());
                cellStyle.setFillPattern(style.fillPattern());
                Font font = createFont();
                font.setFontName(style.fontName());
                font.setFontHeightInPoints(style.fontSize());
                font.setColor(style.fontColor().getIndex());
                cellStyle.setFont(font);
                overrideHeader.put(i, cellStyle);
                if (context.getHeaderHeight() < style.height()) {
                    context.setHeaderHeight(style.height());
                }
            });
            Optional.ofNullable(cs.getBodyStyle()).ifPresent(style -> {
                String format = cs.getFormat();
                CellStyle cellStyle = Optional.ofNullable(context.getBody()).map(cStyle -> {
                    CellStyle override = createCellStyle();
                    override.cloneStyleFrom(cStyle);
                    return override;
                }).orElse(createCellStyle());
                cellStyle.setWrapText(style.isWrap());
                cellStyle.setFillForegroundColor(style.backgroundColor().getIndex());
                cellStyle.setFillPattern(style.fillPattern());
                Font font = createFont();
                font.setFontName(style.fontName());
                font.setFontHeightInPoints(style.fontSize());
                font.setColor(style.fontColor().getIndex());
                cellStyle.setFont(font);
                cellStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat(format));
                overrideBody.put(i, cellStyle);
                if (context.getBodyHeight() < style.height()) {
                    context.setBodyHeight(style.height());
                }
            });
            if (cs.getBodyStyle() == null && !"General".equals(cs.getFormat())) {
                CellStyle cellStyle = createCellStyle();
                cellStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat(cs.getFormat()));
                overrideBody.put(i, cellStyle);
            }
        });
        context.setOverrideHeader(overrideHeader);
        context.setOverrideBody(overrideBody);

        this.context = context;
    }

    private CellStyle createHeaderStyle(HeaderStyle style) {

        // 罫線設定
        CellStyle cellStyle = createCellStyle();
        cellStyle.setBorderBottom(style.borderStyle());
        cellStyle.setBorderLeft(style.borderStyle());
        cellStyle.setBorderRight(style.borderStyle());
        cellStyle.setBorderTop(style.borderStyle());

        // 罫線色設定
        IndexedColors borderColor = style.borderColor();
        cellStyle.setBottomBorderColor(borderColor.getIndex());
        cellStyle.setLeftBorderColor(borderColor.getIndex());
        cellStyle.setRightBorderColor(borderColor.getIndex());
        cellStyle.setTopBorderColor(borderColor.getIndex());

        // 背景色設定
        cellStyle.setFillForegroundColor(style.backgroundColor().getIndex());
        cellStyle.setFillPattern(style.fillPattern());

        // 折り返し設定
        cellStyle.setWrapText(style.isWrap());

        // 水平方向の位置
        cellStyle.setAlignment(style.alignment());

        // フォント設定
        Font font = createFont();
        font.setFontName(style.fontName());
        font.setItalic(style.isItalic());
        font.setBold(style.isBold());
        font.setFontHeightInPoints(style.fontSize());
        font.setColor(style.fontColor().getIndex());
        cellStyle.setFont(font);
        return cellStyle;
    }

    private CellStyle createBodyStyle(BodyStyle style) {

        // 罫線設定
        CellStyle cellStyle = createCellStyle();
        cellStyle.setBorderBottom(style.bodyBorderStyle());
        cellStyle.setBorderLeft(style.borderStyle());
        cellStyle.setBorderRight(style.borderStyle());
        cellStyle.setBorderTop(style.bodyBorderStyle());

        // 罫線色設定
        cellStyle.setBottomBorderColor(style.bodyBorderColor().getIndex());
        cellStyle.setLeftBorderColor(style.borderColor().getIndex());
        cellStyle.setRightBorderColor(style.borderColor().getIndex());
        cellStyle.setTopBorderColor(style.bodyBorderColor().getIndex());

        // 背景色設定
        cellStyle.setFillForegroundColor(style.backgroundColor().getIndex());
        cellStyle.setFillPattern(style.fillPattern());

        // 折り返し設定
        cellStyle.setWrapText(style.isWrap());

        // 水平方向の位置
        cellStyle.setAlignment(style.alignment());

        // font 設定
        Font font = createFont();
        font.setFontName(style.fontName());
        font.setItalic(style.isItalic());
        font.setBold(style.isBold());
        font.setFontHeightInPoints(style.fontSize());
        font.setColor(style.fontColor().getIndex());
        cellStyle.setFont(font);
        return cellStyle;
    }

    private Font createFont() {
        return wb.createFont();
    }

    private CellStyle createCellStyle() {
        return wb.createCellStyle();
    }

    private void writeHeader() {
        String[] headerArray = new String[rowSpec.getLastIndex() + 1];
        rowSpec.getColumns().forEach(field -> {
            Column col = field.getDeclaredAnnotation(Column.class);
            headerArray[col.index()] = col.header();
        });

        Row headerRow = createRow();
        headerRow.setHeight(context.getHeaderHeight());
        int colOffset = spreadSheetSpec.getStartCol();
        IntStream.range(0, headerArray.length)
                .forEach(i -> CellUtil.createCell(headerRow, i + colOffset, headerArray[i],
                        context.getHeaderCellStyle(i)));
    }

    public void writeBody(T body) {

        String[] bodyArray = new String[rowSpec.getLastIndex() + 1];
        rowSpec.getColumns().forEach(field -> {
            Column col = field.getDeclaredAnnotation(Column.class);
            String value = null;
            try {
                PropertyDescriptor pd = new PropertyDescriptor(field.getName(), body.getClass());
                Object fVal = pd.getReadMethod().invoke(body, (Object[]) null);
                value = fVal == null ? null : fVal.toString();
            } catch (IntrospectionException | InvocationTargetException
                    | IllegalAccessException | IllegalArgumentException e) {
                log.warn("Body creation failed. type={}, field={}", rowSpec.getType().getName(), field.getName(), e);
            }
            bodyArray[col.index()] = value;
        });
        List<List<String>> bodies = new ArrayList<>();
        bodies.add(Arrays.asList(bodyArray));

        for (List<String> b : bodies) {
            Row bodyRow = createRow();
            bodyRow.setHeight(context.getBodyHeight());
            String[] array = b.toArray(new String[]{});
            int colOffset = spreadSheetSpec.getStartCol();
            IntStream.range(0, array.length)
                    .forEach(i -> {
                        CellType cellType = rowSpec.getColumnSpecs().get(i).getCellType();
                        Cell cell = bodyRow.createCell(i + colOffset, cellType);
                        setCellValue(cell, cellType, array[i]);
                        cell.setCellStyle(overrideStyleIfFirstRow(context.getBodyCellStyle(i)));
                    });
        }
        bodyIndex++;
    }

    private CellStyle overrideStyleIfFirstRow(CellStyle baseStyle) {
        boolean needsTopBorder = bodyIndex == 0 && context.getBody() != null;
        if (needsTopBorder) {
            return Optional.ofNullable(baseStyle).map(style -> {
                CellStyle newStyle = createCellStyle();
                newStyle.cloneStyleFrom(style);
                newStyle.setBorderTop(context.getBodySurroundStyle());
                newStyle.setTopBorderColor(context.getBodySurroundColor().getIndex());
                return newStyle;
            }).orElse(null);
        } else {
            return baseStyle;
        }
    }

    private void setCellValue(Cell cell, CellType type, String value) {
        if (StringUtils.isEmpty(value)) {
            cell.setCellValue(value);
        } else {
            if (type == CellType.NUMERIC) {
                cell.setCellValue(Double.valueOf(value.replaceAll(",", "")));
            } else {
                cell.setCellValue(new XSSFRichTextString(value));
            }
        }
    }

    private void writeEnd() {
        if (context.getBody() != null) {
            CellStyle cellStyle = createCellStyle();
            cellStyle.setBorderTop(context.getBodySurroundStyle());
            cellStyle.setTopBorderColor(context.getBodySurroundColor().getIndex());

            Row footerRow = createRow();
            IntStream.range(0, this.rowSpec.getLastIndex() + 1)
                    .forEach(i -> {
                        int colOffset = spreadSheetSpec.getStartCol();
                        CellUtil.createCell(footerRow, i + colOffset, null, cellStyle);
                    });
        }
    }

    private Row createRow() {
        return CellUtil.getRow(currentRow++, sheet);
    }

    @Override
    public void close() {
        writeEnd();
        try {
            wb.write(out);
            wb.dispose();
            wb.close();
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    @Data
    private static class StyleContext {
        private CellStyle header;
        private short headerHeight = -1;
        private BorderStyle headerSurroundStyle;
        private IndexedColors headerSurroundColor;
        private CellStyle body;
        private short bodyHeight = -1;
        private BorderStyle bodySurroundStyle;
        private IndexedColors bodySurroundColor;
        private Map<Integer, CellStyle> overrideHeader;
        private Map<Integer, CellStyle> overrideBody;

        private CellStyle getHeaderCellStyle(int i) {
            CellStyle style = header;
            if (overrideHeader.containsKey(i)) {
                style = overrideHeader.get(i);
            }
            return style;
        }

        private CellStyle getBodyCellStyle(int i) {
            CellStyle style = body;
            if (overrideBody.containsKey(i)) {
                style = overrideBody.get(i);
            }
            return style;
        }

    }

}
