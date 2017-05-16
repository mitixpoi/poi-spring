package org.poi.spring;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Created by Hong.LvHang on 2017-05-15.
 */
public class ExcleWapper {
    private Workbook workbook;
    private Sheet sheet;

    public ExcleWapper(Workbook workbook,Sheet sheet) {
        this.workbook = workbook;
        this.sheet = sheet;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }
}
