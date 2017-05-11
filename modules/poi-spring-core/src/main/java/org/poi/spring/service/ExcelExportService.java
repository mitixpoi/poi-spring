package org.poi.spring.service;


import org.apache.poi.ss.usermodel.Sheet;
import org.poi.spring.config.ExcelWorkBookBeandefinition;
import org.poi.spring.service.result.ExcelExportResult;

import java.util.List;

/**
 * Created by Hong.LvHang on 2017-05-08.
 */
public interface ExcelExportService {

    /**
     * 创建Excle表格
     * @param beans
     * @return
     */
    ExcelExportResult createExcel(List<?> beans);

    /**
     * 创建Excle模板表格
     *
     * @param bean
     * @return
     */
    ExcelExportResult createTemplateExcel(Object bean);

    /**
     * 现有表格进行数据追加
     *
     * @param excelWorkBookBeandefinition
     * @param sheet
     * @param beans
     */
    void createRows(ExcelWorkBookBeandefinition excelWorkBookBeandefinition, Sheet sheet, List<?> beans);

}
