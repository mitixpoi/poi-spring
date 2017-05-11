package org.poi.spring.annotation;

import org.poi.spring.PoiConstant;
import org.springframework.stereotype.Component;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by oldflame on 2017/4/8.
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Component
public @interface Excle {

    String name();

    String sheetName() default "";

    short sheetIndex() default 0;

    int width() default PoiConstant.DEFAULT_WIDTH ;








    Align align() default Align.GENERAL;

    short font() default 0;
}
