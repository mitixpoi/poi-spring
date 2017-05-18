package org.poi.spring.component;

import org.poi.spring.PoiConstant;
import org.poi.spring.component.interceptor.TemplateExportInterceptor;
import org.poi.spring.component.interceptor.TemplateImportInterceptor;
import org.poi.spring.config.ExcelWorkBookBeandefinition;
import org.springframework.beans.BeansException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.support.ApplicationObjectSupport;
import org.springframework.core.OrderComparator;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Created by oldflame on 2017/4/8.
 */
public class ExcleContext extends ApplicationObjectSupport {

    private final Set<String> excelIds = new HashSet<>();

    private final Map<String, ExcelWorkBookBeandefinition> excelMap = new HashMap<>();
    /**
     * 默认设置的倒入拦截器
     */
    private final List<TemplateImportInterceptor> templateImportInterceptors = new ArrayList<>();

    /**
     * 默认设置的导出拦截器
     */
    private final List<TemplateExportInterceptor> templateExportInterceptors = new ArrayList<>();
    /**
     * 类型转换器
     */
    @Autowired
    private DefaultExcleConverter excleConverter;

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

        String[] importInterceptors = getApplicationContext().getBeanNamesForType(TemplateImportInterceptor.class);
        if (importInterceptors != null && importInterceptors.length > 0) {
            for (String beanName : importInterceptors) {
                TemplateImportInterceptor templateImportInterceptor = (TemplateImportInterceptor) getApplicationContext().getBean(beanName);
                templateImportInterceptors.add(templateImportInterceptor);
            }
            //排序
            Collections.sort(templateImportInterceptors, OrderComparator.INSTANCE);
        }

        String[] exportInterceptors = getApplicationContext().getBeanNamesForType(TemplateExportInterceptor.class);
        if (exportInterceptors != null && exportInterceptors.length > 0) {
            for (String beanName : exportInterceptors) {
                TemplateExportInterceptor templateExportInterceptor = (TemplateExportInterceptor) getApplicationContext().getBean(beanName);
                templateExportInterceptors.add(templateExportInterceptor);
            }
            //排序
            Collections.sort(templateExportInterceptors, OrderComparator.INSTANCE);
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

    public DefaultExcleConverter getExcleConverter() {
        return excleConverter;
    }

    public void setExcleConverter(DefaultExcleConverter excleConverter) {
        this.excleConverter = excleConverter;
    }

    public List<TemplateImportInterceptor> getTemplateImportInterceptors() {
        return templateImportInterceptors;
    }

    public List<TemplateExportInterceptor> getTemplateExportInterceptors() {
        return templateExportInterceptors;
    }
}
