package naganaga.ss.annotations;

import org.apache.poi.ss.usermodel.CellType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD })
public @interface Column {
    int index();
    int width() default -1;
    String header() default "";
    CellType cellType() default CellType.STRING;
    String format() default "General";
}
