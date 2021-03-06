package org.poi.spring;

import org.apache.poi.hssf.util.CellReference;
import org.poi.spring.config.ColumnDefinition;

/**
 * Created by Hong.LvHang on 2017-05-09.
 */
public class CellReferenceWapper {
    //cell ref
    private CellReference cellReference;
    //cell definition
    private ColumnDefinition columnDefinition;
    //cell interceptor errMessage
    private String errMessage;

    public CellReferenceWapper(CellReference cellReference, ColumnDefinition columnDefinition) {
        this.columnDefinition = columnDefinition;
        this.cellReference = cellReference;
    }

    public CellReference getCellReference() {
        return cellReference;
    }

    public void setCellReference(CellReference cellReference) {
        this.cellReference = cellReference;
    }

    public ColumnDefinition getColumnDefinition() {
        return columnDefinition;
    }

    public void setColumnDefinition(ColumnDefinition columnDefinition) {
        this.columnDefinition = columnDefinition;
    }

    public String getErrMessage() {
        return errMessage;
    }

    public void setErrMessage(String errMessage) {
        this.errMessage = errMessage;
    }
}
