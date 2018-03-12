package naganaga.ss.spec;

import naganaga.ss.annotations.BodyStyle;
import naganaga.ss.annotations.HeaderStyle;
import naganaga.ss.annotations.SpreadSheet;

/**
 * スプレッドシートの仕様に関するクラス.
 *
 */
public class SpreadSheetSpec {

    /** シート名. */
    private String sheetName;
    /** 開始行番号. */
    private int startRow;
    /** 開始列番号. */
    private int startCol;
    /** ヘッダ有か否か. */
    private boolean hasHeader;
    private HeaderStyle headerStyle;
    private BodyStyle bodyStyle;

    /**
     * コンストラクタ.
     *
     * @param type 型
     */
    public SpreadSheetSpec(Class<?> type) {
        SpreadSheet format = type.getDeclaredAnnotation(SpreadSheet.class);
        if (format == null) {
            // @SpreadSheetFormat が指定されていることをチェック
            throw new IllegalArgumentException("@SpreadSheetFormat not defined. type=[" + type.getName() + "]");
        }
        this.hasHeader = format.writeHeader();
        this.sheetName = format.name();
        this.startRow = format.startRowNumber();
        this.startCol = format.startColumnNumber();
        this.headerStyle = type.getDeclaredAnnotation(HeaderStyle.class);
        this.bodyStyle = type.getDeclaredAnnotation(BodyStyle.class);
    }

    public String getSheetName() {
        return sheetName;
    }

    public int getStartRow() {
        return startRow;
    }

    public int getStartCol() {
        return startCol;
    }

    public boolean hasHeader() {
        return hasHeader;
    }

    public HeaderStyle getHeaderStyle() {
        return headerStyle;
    }

    public BodyStyle getBodyStyle() {
        return bodyStyle;
    }

}
