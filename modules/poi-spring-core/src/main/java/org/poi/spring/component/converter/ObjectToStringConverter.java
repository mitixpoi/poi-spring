package org.poi.spring.component.converter;

import org.springframework.core.convert.converter.Converter;

/**
 * Created by Hong.LvHang on 2017-05-16.
 */
public final class ObjectToStringConverter implements Converter<Object, String> {

    public String convert(Object source) {
        return source.toString();
    }
}
