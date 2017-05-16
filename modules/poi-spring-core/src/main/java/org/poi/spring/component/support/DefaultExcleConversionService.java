package org.poi.spring.component.support;

import org.springframework.core.convert.ConversionService;
import org.springframework.core.convert.converter.ConverterRegistry;
import org.springframework.core.convert.support.DefaultConversionService;
import org.springframework.core.convert.support.GenericConversionService;

/**
 * 类型转换
 * Created by Hong.LvHang on 2017-05-08.
 */

public class DefaultExcleConversionService extends GenericConversionService {

    private static volatile DefaultExcleConversionService sharedInstance;


    public static ConversionService getSharedInstance() {
        if (sharedInstance == null) {
            synchronized (DefaultConversionService.class) {
                if (sharedInstance == null) {
                    sharedInstance = new DefaultExcleConversionService();
                }
            }
        }
        return sharedInstance;
    }


    public DefaultExcleConversionService() {
        addDefaultConverters(this);
    }

    private void addDefaultConverters(ConverterRegistry converterRegistry) {
        converterRegistry.addConverterFactory(new NumberToNumberConverterFactory());

        converterRegistry.addConverterFactory(new StringToNumberConverterFactory());
        converterRegistry.addConverter(Number.class, String.class, new ObjectToStringConverter());

        converterRegistry.addConverter(new StringToBooleanConverter());
        converterRegistry.addConverter(Boolean.class, String.class, new ObjectToStringConverter());

        converterRegistry.addConverter(new DateToStringConverter());
        converterRegistry.addConverter(new StringToDateConverter());

    }
}
