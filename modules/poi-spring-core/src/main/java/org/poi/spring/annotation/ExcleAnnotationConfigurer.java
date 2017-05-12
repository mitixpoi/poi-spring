package org.poi.spring.annotation;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellUtil;
import org.poi.spring.PoiConstant;
import org.poi.spring.config.ColumnDefinition;
import org.poi.spring.config.ExcelWorkBookBeandefinition;
import org.poi.spring.exception.ExcelException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.BeansException;
import org.springframework.beans.factory.config.BeanDefinition;
import org.springframework.beans.factory.config.ConfigurableListableBeanFactory;
import org.springframework.beans.factory.support.BeanDefinitionRegistry;
import org.springframework.beans.factory.support.BeanDefinitionRegistryPostProcessor;
import org.springframework.beans.factory.support.RootBeanDefinition;
import org.springframework.core.annotation.AnnotationUtils;
import org.springframework.util.ClassUtils;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Hong.LvHang on 2017-04-12.
 */
public class ExcleAnnotationConfigurer implements BeanDefinitionRegistryPostProcessor {
    private static final Logger logger = LoggerFactory.getLogger(ExcleAnnotationConfigurer.class);


    public void postProcessBeanDefinitionRegistry(BeanDefinitionRegistry beanDefinitionRegistry) throws BeansException {
        String[] candidateBean = beanDefinitionRegistry.getBeanDefinitionNames();
        for (String beanName : candidateBean) {
            BeanDefinition beanDefinition = beanDefinitionRegistry.getBeanDefinition(beanName);

            String originClassName = beanDefinition.getBeanClassName();
            Class originClass = null;
            try {
                originClass = ClassUtils.forName(originClassName, null);
            } catch (ClassNotFoundException e) {
                logger.error("excleworkbook class not fand originClassName=" + originClassName);
                ExcelException exception = new ExcelException("excleworkbook class not fand originClassName=" + originClassName);
                exception.initCause(e);
                throw exception;
            }
            Annotation annotation = AnnotationUtils.getAnnotation(originClass, Excle.class);
            if (annotation != null) {
                String id = originClass.getName() + PoiConstant.CLASS_NAME_SUFFIX;
                RootBeanDefinition eligibleBeanDefinition = new RootBeanDefinition();
                eligibleBeanDefinition.setBeanClass(ExcelWorkBookBeandefinition.class);
                eligibleBeanDefinition.setLazyInit(false);
                eligibleBeanDefinition.getPropertyValues().addPropertyValue("dataClass", originClass);
                eligibleBeanDefinition.getPropertyValues().addPropertyValue("id", id);
                Map<String, Object> attributesMap = AnnotationUtils.getAnnotationAttributes(annotation);
                //excleName
                if (attributesMap.get("name") == null) {
                    throw new ExcelException("Annotation Excle name not fand originClass=" + originClass.getName());
                }
                eligibleBeanDefinition.getPropertyValues().addPropertyValue("excleName", attributesMap.get("name"));
                if (attributesMap.get("sheetName") != null) {
                    eligibleBeanDefinition.getPropertyValues().addPropertyValue("sheetName", attributesMap.get("sheetName"));
                }
                if (attributesMap.get("sheetIndex") != null) {
                    eligibleBeanDefinition.getPropertyValues().addPropertyValue("sheetIndex", attributesMap.get("sheetIndex"));
                }
                if (attributesMap.get("width") != null) {
                    try {
                        int width = Integer.valueOf(String.valueOf(attributesMap.get("width"))) * PoiConstant.DEFAULT_WIDTH_R;
                        eligibleBeanDefinition.getPropertyValues().addPropertyValue("columnWidth", width);
                    } catch (Exception e) {
                        ExcelException exception = new ExcelException("设置宽度失败 excleName=" + attributesMap.get("name"));
                        exception.initCause(e);
                        throw exception;
                    }
                }
                //参数信息
                Map<String, Object> defaultProperties = addDefaultProperties(attributesMap);
                eligibleBeanDefinition.getPropertyValues().addPropertyValue("defaultProperties", defaultProperties);
                //解析属性上的注解
                List<ColumnDefinition> columnDefinitions = parseColumnAnnotations(originClass);
                eligibleBeanDefinition.getPropertyValues().addPropertyValue("columnDefinitions", columnDefinitions);
                beanDefinitionRegistry.registerBeanDefinition(id, eligibleBeanDefinition);
            }
        }
    }

    private List<ColumnDefinition> parseColumnAnnotations(Class originClass) {
        List<ColumnDefinition> columnDefinitions = new ArrayList<>();
        Field[] fields = originClass.getDeclaredFields();
        for (Field field : fields) {
            Annotation annotation = AnnotationUtils.getAnnotation(field, Column.class);
            if (annotation != null) {
                Map<String, Object> attributesMap = AnnotationUtils.getAnnotationAttributes(annotation);
                columnDefinitions.add(parseColumnAnnotation(field, attributesMap));
            }
        }
        return columnDefinitions;
    }

    private ColumnDefinition parseColumnAnnotation(Field field, Map<String, Object> attributesMap) {
        ColumnDefinition columnDefinition = new ColumnDefinition();
        columnDefinition.setName(field.getName());
        columnDefinition.setTitle(String.valueOf(attributesMap.get("title")));
        columnDefinition.setRegex(String.valueOf(attributesMap.get("regex")));
        columnDefinition.setRequired((Boolean) attributesMap.get("required"));
        columnDefinition.setColumnWidth((Integer) attributesMap.get("width") * PoiConstant.DEFAULT_WIDTH_R);
        columnDefinition.setFormat(String.valueOf(attributesMap.get(PoiConstant.FORMAT)));
        columnDefinition.setDictNo(String.valueOf(attributesMap.get(PoiConstant.DICTNO)));
        //        if (!PoiConstant.EMPTY_STRING.equals(attributesMap.get("defauleValue"))) {
        //            columnDefinition.setDefaultValue(String.valueOf(attributesMap.get("defauleValue")));
        //        }
        //参数信息
        Map<String, Object> properties = addProperties(attributesMap);
        columnDefinition.setProperties(properties);
        return columnDefinition;
    }

    private Map<String, Object> addProperties(Map<String, Object> attributesMap) {
        Map<String, Object> properties = new HashMap<>();
        addfgcolorProperties(properties, attributesMap.get("fgcolor"));
        addBorderProperties(properties, attributesMap.get("border"));
        addAlignProperties(properties, attributesMap.get("align"));

        //        addFontProperties(properties, attributesMap.get("font"));
        //        addWrapTextProperties(properties, attributesMap.get("wraptext"));
        return properties;
    }

    private Map<String, Object> addDefaultProperties(Map<String, Object> attributesMap) {
        Map<String, Object> defaultProperties = new HashMap<>();
        addfgcolorProperties(defaultProperties, attributesMap.get("fgcolor"));
        addBorderProperties(defaultProperties, attributesMap.get("border"));
        addAlignProperties(defaultProperties, attributesMap.get("align"));
        //        addFontProperties(defaultProperties, attributesMap.get("font"));
        return defaultProperties;
    }

    private void addBorderProperties(Map<String, Object> properties, Object border) {
        if (border instanceof BorderStyle && !BorderStyle.NONE.equals(border)) {
            //设置边线样式
            properties.put(CellUtil.BORDER_TOP, border);
            properties.put(CellUtil.BORDER_BOTTOM, border);
            properties.put(CellUtil.BORDER_LEFT, border);
            properties.put(CellUtil.BORDER_RIGHT, border);
        }
    }

    private void addfgcolorProperties(Map<String, Object> properties, Object fgcolor) {
        if (fgcolor instanceof IndexedColors && !IndexedColors.AUTOMATIC.equals(fgcolor)) {
            properties.put(CellUtil.FILL_FOREGROUND_COLOR, ((IndexedColors) fgcolor).getIndex());
            properties.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
        }
    }

    private void addWrapTextProperties(Map<String, Object> properties, Object wraptext) {
        if (wraptext instanceof Boolean) {
            properties.put(CellUtil.WRAP_TEXT, wraptext);
        }
    }

    private void addFontProperties(Map<String, Object> properties, Object font) {
        if (font instanceof Short) {
            properties.put(CellUtil.FONT, font);
        }
    }

    private void addAlignProperties(Map<String, Object> properties, Object align) {
        if (align instanceof HorizontalAlignment && !HorizontalAlignment.GENERAL.equals(align)) {
            properties.put(CellUtil.ALIGNMENT, align);
        }
    }

    public void postProcessBeanFactory(ConfigurableListableBeanFactory configurableListableBeanFactory) throws BeansException {
        //无需处理
    }
}
