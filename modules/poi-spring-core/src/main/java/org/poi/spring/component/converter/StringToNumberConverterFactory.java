package org.poi.spring.component.converter;

import org.springframework.core.convert.converter.Converter;
import org.springframework.core.convert.converter.ConverterFactory;
import org.springframework.util.NumberUtils;

/**
 * Created by Hong.LvHang on 2017-05-16.
 */
public class StringToNumberConverterFactory implements ConverterFactory<String, Number> {

    public <T extends Number> Converter<String, T> getConverter(Class<T> targetType) {
        return new StringToNumberConverterFactory.StringToNumber(targetType);
    }

    private static final class StringToNumber<T extends Number> implements Converter<String, T> {
        private final Class<T> targetType;

        public StringToNumber(Class<T> targetType) {
            this.targetType = targetType;
        }

        public T convert(String source) {
            return source.length() == 0 ? null : NumberUtils.parseNumber(source, this.targetType);
        }
    }
}
