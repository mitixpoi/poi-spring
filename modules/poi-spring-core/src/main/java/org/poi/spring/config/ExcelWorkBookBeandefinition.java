package org.poi.spring.config;

import org.poi.spring.Assert;
import org.poi.spring.PoiConstant;
import org.springframework.beans.factory.InitializingBean;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Created by oldflame on 2017/4/6.
 */
public class ExcelWorkBookBeandefinition implements InitializingBean {

    /**
     * ID,必须（类名+.EXCLE）
     */
    private String id;

    /**
     * 全类名,必须(导出名字)
     */
    private String excleName;

    /**
     * Class信息
     */
    private Class<?> dataClass;

    /**
     * 导出时,sheet页的名称,可以不设置
     */
    private String sheetName;

    /**
     * Excel 文件sheet索引，默认为0即，第一个
     */
    private int sheetIndex = 0;

    private String header;

    /**
     * 实际使用宽度
     */
    private Integer columnWidth;

    /**
     * 参数信息
     */
    private Map<String, Object> defaultProperties;

    /**
     * Field属性的全部定义
     */
    private List<ColumnDefinition> columnDefinitions = new ArrayList<>();


    public Integer getColumnWidth() {
        return columnWidth;
    }

    public void setColumnWidth(Integer columnWidth) {
        this.columnWidth = columnWidth;
    }

    public Map<String, Object> getDefaultProperties() {
        return defaultProperties;
    }

    public void setDefaultProperties(Map<String, Object> defaultProperties) {
        this.defaultProperties = defaultProperties;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<ColumnDefinition> getColumnDefinitions() {
        return columnDefinitions;
    }

    public void setColumnDefinitions(List<ColumnDefinition> columnDefinitions) {
        this.columnDefinitions = columnDefinitions;
    }

    public Class<?> getDataClass() {
        return dataClass;
    }

    public void setDataClass(Class<?> dataClass) {
        this.dataClass = dataClass;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getExcleName() {
        return excleName;
    }

    public void setExcleName(String excleName) {
        this.excleName = excleName;
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public String getHeader() {
        return header;
    }

    public void setHeader(String header) {
        this.header = header;
    }

    @Override
    public void afterPropertiesSet() throws Exception {
        //todo  可以做一些有效性的判断
        Assert.hasNoColumn(this.excleName, this.columnDefinitions);
        setDefaultFields();
        mergeProperties(this.getDefaultProperties(), this.getColumnDefinitions());
    }

    private void setDefaultFields() {
        //设置默认的属性
        if (this.getSheetName() == null || "".equals(this.getSheetName())) {
            this.setSheetName("sheet" + this.sheetIndex);
        }

    }

    //进行属性合并
    private void mergeProperties(Map<String, Object> defaultProperties, List<ColumnDefinition> columnDefinitions) {
        for (ColumnDefinition columnDefinition : columnDefinitions) {
            //设置字段长度
            if (columnDefinition.getColumnWidth() == null) {
                columnDefinition.setColumnWidth(
                    this.getColumnWidth() == null ? PoiConstant.DEFAULT_WIDTH * PoiConstant.DEFAULT_WIDTH_R : this.getColumnWidth());
            }
            Map<String, Object> properties = columnDefinition.getProperties();
            //进行合并   JDK1.8的Map有merge方法  这里不进行使用
            for (String name : defaultProperties.keySet()) {
                if (properties.get(name) == null) {
                    properties.put(name, defaultProperties.get(name));
                }
            }
        }
    }
}
