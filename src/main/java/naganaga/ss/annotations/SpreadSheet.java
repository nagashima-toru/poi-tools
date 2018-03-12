package naganaga.ss.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.TYPE })
public @interface SpreadSheet {
    String name() default "Sheet1";
    boolean writeHeader() default true;
    int startRowNumber() default 0;
    int startColumnNumber() default 0;
}
