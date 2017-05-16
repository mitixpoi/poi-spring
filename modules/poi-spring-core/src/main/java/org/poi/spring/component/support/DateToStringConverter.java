package org.poi.spring.component.support;

import org.springframework.core.convert.converter.Converter;

import java.util.Date;

/**
 * Created by Hong.LvHang on 2017-05-16.
 */
public class DateToStringConverter implements Converter<Date, String> {

    @Override
    public String convert(Date source) {
        //todo 日期转换器实现
        return null;
    }
}
