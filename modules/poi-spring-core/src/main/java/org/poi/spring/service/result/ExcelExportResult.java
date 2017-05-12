package org.poi.spring.service.result;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.poi.spring.config.ExcelWorkBookBeandefinition;

/**
 * Excel导出结果
 */
public class ExcelExportResult {
    private ExcelWorkBookBeandefinition excelWorkBookBeandefinition;
    private Sheet sheet;
    private Workbook workbook;

    public Workbook getWorkbook() {
        return workbook;
    }

    public ExcelExportResult(ExcelWorkBookBeandefinition excelDefinition, Sheet sheet, Workbook workbook) {
        super();
        this.excelWorkBookBeandefinition = excelDefinition;
        this.sheet = sheet;
        this.workbook = workbook;
    }


    /**
     * 导出完毕,获取WorkBook
     *
     * @return
     */
    public Workbook build() {
        return workbook;
    }

}
