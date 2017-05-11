package org.poi.spring.component;

import org.apache.poi.ss.usermodel.Sheet;
import org.poi.spring.config.ExcelWorkBookBeandefinition;

import java.util.List;
/**
 * Created by Hong.LvHang on 2017-05-11.
 */
public interface ReportExcleHeader extends ExcelHeader {
    /**
     * 如何构建标题之前的数据
     *
     * @param sheet           Excel中的sheet页
     * @param excelDefinition XML中定义的信息
     * @param beans           导出的数据
     */
    void buildHeader(Sheet sheet, ExcelWorkBookBeandefinition excelDefinition, List<?> beans);
}
