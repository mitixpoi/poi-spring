package org.poi.spring.config;

import java.util.List;
import java.util.Map;

/**
 * Created by oldflame on 2017/4/6.
 */
public class ColumnDefinition {
    /**
     * 属性名称,必须
     */
    private String name;
    /**
     * 标题,必须
     */
    private String title;
    /**
     * 是否必须
     */
    private boolean required = false;

    /**
     * 导出拦截器
     */
    private List<Object> templateExportInterceptors;
    /**
     * 导入拦截器
     */
    private List<Object> templateImportInterceptors;

    /**
     * 表达式
     */
    private String regex;

    /**
     * 校验失败返回
     */
    private String validatorErrMsg;

    //导出时生效
    /**
     * cell的宽度,此属性不受enableStyle影响,如自己把数值写的过大
     *
     * @see org.apache.poi.ss.usermodel.Sheet
     */
    private int columnWidth;

    /**
     * 当值为空时,字段的默认值
     */
    private String defaultValue;

    /**
     * 表达式,例如(1:男,2:女)表示,值为1,取 (男)作为value ,2则取 (女)作为value
     */
    private String format;

    /**
     * 字典值，取数据库
     */
    private String dictNo;

    /**
     * 参数信息
     */
    private Map<String, Object> properties;

    public List<Object> getTemplateExportInterceptors() {
        return templateExportInterceptors;
    }

    public void setTemplateExportInterceptors(List<Object> templateExportInterceptors) {
        this.templateExportInterceptors = templateExportInterceptors;
    }

    public List<Object> getTemplateImportInterceptors() {
        return templateImportInterceptors;
    }

    public void setTemplateImportInterceptors(List<Object> templateImportInterceptors) {
        this.templateImportInterceptors = templateImportInterceptors;
    }

    public void setColumnWidth(int columnWidth) {
        this.columnWidth = columnWidth;
    }

    public Map<String, Object> getProperties() {
        return properties;
    }

    public void setProperties(Map<String, Object> properties) {
        this.properties = properties;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public boolean isRequired() {
        return required;
    }

    public void setRequired(boolean required) {
        this.required = required;
    }

    public String getRegex() {
        return regex;
    }

    public void setRegex(String regex) {
        this.regex = regex;
    }

    public String getValidatorErrMsg() {
        return validatorErrMsg;
    }

    public void setValidatorErrMsg(String validatorErrMsg) {
        this.validatorErrMsg = validatorErrMsg;
    }

    public Integer getColumnWidth() {
        return columnWidth;
    }

    public void setColumnWidth(Integer columnWidth) {
        this.columnWidth = columnWidth;
    }

    public String getDefaultValue() {
        return defaultValue;
    }

    public void setDefaultValue(String defaultValue) {
        this.defaultValue = defaultValue;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public String getDictNo() {
        return dictNo;
    }

    public void setDictNo(String dictNo) {
        this.dictNo = dictNo;
    }
}
