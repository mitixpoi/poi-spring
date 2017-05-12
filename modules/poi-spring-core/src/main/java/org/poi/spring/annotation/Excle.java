package org.poi.spring.annotation;

import com.sun.xml.internal.ws.api.streaming.XMLStreamReaderFactory;
import org.apache.poi.ss.usermodel.BorderStyle;
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

    int width() default PoiConstant.DEFAULT_WIDTH ;
    //todo 属性在xml中还未处理
    IndexedColors fgcolor() default IndexedColors.AUTOMATIC;
    //todo 属性在xml中还未处理
    BorderStyle border() default BorderStyle.NONE;







    Align align() default Align.GENERAL;

    short font() default 0;
}
