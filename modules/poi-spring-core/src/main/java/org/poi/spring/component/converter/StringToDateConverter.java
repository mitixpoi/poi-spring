package org.poi.spring.component.converter;

import org.springframework.core.convert.converter.Converter;

import java.util.Date;

/**
 * Created by Hong.LvHang on 2017-05-16.
 */
public class StringToDateConverter implements Converter<String, Date> {

    @Override
    public Date convert(String source) {
        //todo 日期转换器实现
        return null;
    }
}
