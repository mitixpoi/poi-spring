package org.poi.spring.service;

import org.poi.spring.service.result.ExcelExportResult;

/**
 * Created by Hong.LvHang on 2017-05-12.
 */
public interface ExcleTemplateExportService {

    /**
     * 创建Excle模板表格
     *
     * @param bean
     * @return
     */
    ExcelExportResult createTemplateExcel(Object bean);
}
