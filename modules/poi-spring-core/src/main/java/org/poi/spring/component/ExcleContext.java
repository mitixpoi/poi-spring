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
    /**
     * 默认设置的倒入拦截器
     */
    private final Map<String, TemplateImportInterceptor> templateImportInterceptorMap = new HashMap<>();
    /**
     * 默认设置的导出拦截器
     */
    private final Map<String, TemplateExportInterceptor> templateExportInterceptorMap = new HashMap<>();

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

        String[] templateImportInterceptors = getApplicationContext().getBeanNamesForType(TemplateImportInterceptor.class);
        if (templateImportInterceptors != null && templateImportInterceptors.length > 0) {
            for (String beanName : templateImportInterceptors) {
                TemplateImportInterceptor templateImportInterceptor = (TemplateImportInterceptor) getApplicationContext().getBean(beanName);
                templateImportInterceptorMap.put(beanName, templateImportInterceptor);
            }
        }

        String[] templateExportInterceptors = getApplicationContext().getBeanNamesForType(TemplateExportInterceptor.class);
        if (templateExportInterceptors != null && templateExportInterceptors.length > 0) {
            for (String beanName : templateExportInterceptors) {
                TemplateExportInterceptor templateExportInterceptor = (TemplateExportInterceptor) getApplicationContext().getBean(beanName);
                templateExportInterceptorMap.put(beanName, templateExportInterceptor);
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

    public Map<String, TemplateImportInterceptor> getTemplateImportInterceptorMap() {
        return templateImportInterceptorMap;
    }

    public Map<String, TemplateExportInterceptor> getTemplateExportInterceptorMap() {
        return templateExportInterceptorMap;
    }
}
