package org.poi.spring.annotation;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.poi.spring.PoiConstant;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by Hong.LvHang on 2017-04-13.
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Column {

    String title();

    int width() default PoiConstant.DEFAULT_WIDTH;
    //todo 属性在xml中还未处理
    IndexedColors fgcolor() default IndexedColors.AUTOMATIC;
    //todo 属性在xml中还未处理
    BorderStyle border() default BorderStyle.NONE;

    HorizontalAlignment align() default HorizontalAlignment.GENERAL;

    //格式化约定字典
    String format() default PoiConstant.EMPTY_STRING;
    //数据库字典值
    String dictNo() default PoiConstant.EMPTY_STRING;



    boolean required() default false;

    String regex() default PoiConstant.EMPTY_STRING;


    String defauleValue() default "";
}
