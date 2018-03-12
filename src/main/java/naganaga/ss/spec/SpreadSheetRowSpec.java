package naganaga.ss.spec;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import naganaga.ss.annotations.BodyStyle;
import naganaga.ss.annotations.Column;
import naganaga.ss.annotations.HeaderStyle;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.CellType;

import java.lang.reflect.Field;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * 行に対する設定情報を保持するクラス。
 *
 * @param <T> 変換定義が設定されている型
 */
@Slf4j
public class SpreadSheetRowSpec<T> {

    /** 型. */
    private Class<T> type;
    /** 列（整列済）. */
    private List<ColumnSpec> colSpecs;
    /** 終端列番号 */
    private int lastIndex;

    /**
     * コンストラクタ。
     * @param rowClass 行の情報が設定されたクラス
     */
    public SpreadSheetRowSpec(Class<T> rowClass) {

        this.type = rowClass;
        List<Field> columns = FieldUtils.getFieldsListWithAnnotation(this.type, Column.class);

        if (columns.isEmpty()) {
            // @Column (1個以上) が指定されていることをチェック
            throw new IllegalArgumentException("@Column not defined. type=[" + type.getName() + "]");
        }

        // index 順にソート
        columns.sort((c1, c2) -> {
            Column a1 = c1.getDeclaredAnnotation(Column.class);
            Column a2 = c2.getDeclaredAnnotation(Column.class);
            return Integer.compare(a1.index(), a2.index());
        });

        verifyColumns(columns);

        this.lastIndex = columns.get(columns.size() - 1).getDeclaredAnnotation(Column.class).index();
        this.colSpecs = columns.stream().map(field -> {
            Column spec = field.getDeclaredAnnotation(Column.class);
            ColumnSpec columnSpec = new ColumnSpec();
            columnSpec.setCellType(spec.cellType());
            columnSpec.setWidth(spec.width());
            columnSpec.setField(field);
            columnSpec.setIndex(spec.index());
            columnSpec.setHeader(spec.header());
            columnSpec.setFormat(spec.format());
            columnSpec.setHeaderStyle(field.getDeclaredAnnotation(HeaderStyle.class));
            columnSpec.setBodyStyle(field.getDeclaredAnnotation(BodyStyle.class));
            return columnSpec;
        }).collect(Collectors.toList());
    }

    /**
     * 元となる型を取得する。
     *
     * @return 元となる型
     */
    public Class<T> getType() {
        return type;
    }

    /**
     * {@link Column} が定義されたフィールド一覧を取得する。
     * <pre>
     *     呼出し元でリストの操作をされると困るので、変更不可能なリストを返します。
     * </pre>
     *
     * @return {@link Column} が定義されたフィールド一覧
     */
    public List<Field> getColumns() {
        return Collections.unmodifiableList(
                this.colSpecs.stream().map(ColumnSpec::getField).collect(Collectors.toList()));
    }

    /**
     * {@link Column} が定義されたフィールド一覧を取得する。
     * <pre>
     *     呼出し元でリストの操作をされると困るので、変更不可能なリストを返します。
     * </pre>
     *
     * @return {@link Column} が定義されたフィールド一覧
     */
    public List<ColumnSpec> getColumnSpecs() {
        return this.colSpecs;
    }

    /**
     * 最終列番号を取得する。
     *
     * @return 最終列番号
     */
    public int getLastIndex() {
        return lastIndex;
    }

    private void verifyColumns(List<Field> columns) {
        IntStream.range(0, columns.size()).forEach(i -> {
            Field field = columns.get(i);
            Class<?> fieldType = field.getType();
            // String 型以外はサポートしないので、String であることをチェック
            if (fieldType != String.class) {
                throw new IllegalArgumentException("An unsupported type was specified. type=[" + fieldType.getName() + "]");
            }

            // index に飛び番、重複などがある場合は警告ログを出力する
            Column col = field.getDeclaredAnnotation(Column.class);
            if (i != col.index()) {
                log.warn("The index is not sequential number. type={}, field={}, expected={}, actual={}",
                        this.type.getName(), field.getName(), i, col.index());
            }
        });
    }

    @Data
    public static class ColumnSpec {
        private Field field;
        private CellType cellType;
        private int index;
        private int width;
        private String header;
        private String format;
        private HeaderStyle headerStyle;
        private BodyStyle bodyStyle;
    }

}
