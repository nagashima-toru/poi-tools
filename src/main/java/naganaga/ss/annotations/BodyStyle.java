package naganaga.ss.annotations;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.TYPE, ElementType.FIELD })
public @interface BodyStyle {
    String fontName() default "ＭＳ ゴシック";
    short fontSize() default 11;
    IndexedColors fontColor() default IndexedColors.BLACK;
    IndexedColors backgroundColor() default IndexedColors.WHITE;
    FillPatternType fillPattern() default FillPatternType.NO_FILL;
    boolean isBold() default false;
    boolean isItalic() default false;
    boolean isWrap() default false;
    BorderStyle borderStyle() default BorderStyle.THIN;
    IndexedColors borderColor() default IndexedColors.BLACK;
    BorderStyle bodyBorderStyle() default BorderStyle.THIN;
    IndexedColors bodyBorderColor() default IndexedColors.BLACK;
    short height() default -1;
    HorizontalAlignment alignment() default HorizontalAlignment.GENERAL;
}
