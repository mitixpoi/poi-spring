package org.poi.spring.component;

import org.poi.spring.component.converter.DefaultExcleConversionService;
import org.springframework.core.convert.ConversionService;
import org.springframework.core.convert.TypeDescriptor;

/**
 * 类型转换
 * Created by Hong.LvHang on 2017-05-08.
 */

public class DefaultExcleConverter {

    private final ConversionService conversionService = DefaultExcleConversionService.getSharedInstance();

    public boolean canConvertString(Class<?> sourceType) {
        return canConvert(TypeDescriptor.valueOf(sourceType), TypeDescriptor.valueOf(String.class));
    }

    public boolean canConvert(TypeDescriptor sourceType, TypeDescriptor targetType) {
        return this.conversionService.canConvert(sourceType, targetType);
    }

    public <T> T convert(Object source, Class<T> targetType) {
        return this.conversionService.convert(source, targetType);
    }
}
