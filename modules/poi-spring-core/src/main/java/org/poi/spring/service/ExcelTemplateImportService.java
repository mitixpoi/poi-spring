package org.poi.spring.service;

import org.poi.spring.service.result.ExcelImportResult;

import java.io.InputStream;

/**
 * Created by Hong.LvHang on 2017-05-08.
 */
public interface ExcelTemplateImportService {
    ExcelImportResult readExcel(Class<?> beanClass, InputStream excelStream);
}
