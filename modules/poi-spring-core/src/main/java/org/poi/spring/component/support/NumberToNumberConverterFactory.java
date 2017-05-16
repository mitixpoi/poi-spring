package org.poi.spring.component.support;

import org.springframework.core.convert.converter.Converter;
import org.springframework.core.convert.converter.ConverterFactory;
import org.springframework.util.NumberUtils;

/**
 * Created by Hong.LvHang on 2017-05-16.
 */
public class NumberToNumberConverterFactory<T extends Number> implements ConverterFactory<Number, Number> {

    public <T extends Number> Converter<Number, T> getConverter(Class<T> targetType) {
        return new NumberToNumberConverterFactory.NumberToNumber(targetType);
    }


    private static final class NumberToNumber<T extends Number> implements Converter<Number, T> {
        private final Class<T> targetType;

        public NumberToNumber(Class<T> targetType) {
            this.targetType = targetType;
        }

        public T convert(Number source) {
            return NumberUtils.convertNumberToTargetClass(source, this.targetType);
        }
    }
}
