package org.poi.spring.component;

import org.poi.spring.PoiConstant;
import org.poi.spring.config.ExcelWorkBookBeandefinition;
import org.springframework.beans.BeansException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.support.ApplicationObjectSupport;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

/**
 * Created by oldflame on 2017/4/8.
 */
public class ExcleContext extends ApplicationObjectSupport {

    private final Set<String> excelIds = new HashSet<>();

    private final Map<String, ExcelWorkBookBeandefinition> excelMap = new HashMap<>();

    @Autowired
    private ExcleConverter excleConverter;

    @Autowired(required = false)
    private ExcelDictService excelDictService;

    @Override
    protected void initApplicationContext() throws BeansException {
        String[] excelBeans = getApplicationContext().getBeanNamesForType(ExcelWorkBookBeandefinition.class);
        if (excelBeans != null && excelBeans.length > 0) {
            for (String beanName : excelBeans) {
                ExcelWorkBookBeandefinition excelDef = (ExcelWorkBookBeandefinition) getApplicationContext().getBean(beanName);
                excelIds.add(excelDef.getId());
                excelMap.put(excelDef.getId(), excelDef);
            }
        }
    }

    public boolean exists(Class<?> targetClass) {
        return excelIds.contains(targetClass.getName() + PoiConstant.CLASS_NAME_SUFFIX);
    }

    public ExcelWorkBookBeandefinition getExcelWorkBookBeandefinition(Class<?> targetClass) {
        return excelMap.get(targetClass.getName() + PoiConstant.CLASS_NAME_SUFFIX);
    }

    public ExcelDictService getExcelDictService() {
        return excelDictService;
    }

    public void setExcelDictService(ExcelDictService excelDictService) {
        this.excelDictService = excelDictService;
    }

    public ExcleConverter getExcleConverter() {
        return excleConverter;
    }

    public void setExcleConverter(ExcleConverter excleConverter) {
        this.excleConverter = excleConverter;
    }
}
