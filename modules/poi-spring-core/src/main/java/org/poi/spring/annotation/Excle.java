package org.poi.spring.annotation;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
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

    //设置表头信息
    String header() default "";

    int width() default PoiConstant.DEFAULT_WIDTH ;
    //todo 属性在xml中还未处理
    IndexedColors fgcolor() default IndexedColors.AUTOMATIC;
    //todo 属性在xml中还未处理
    BorderStyle border() default BorderStyle.NONE;

    HorizontalAlignment align() default HorizontalAlignment.GENERAL;


}
