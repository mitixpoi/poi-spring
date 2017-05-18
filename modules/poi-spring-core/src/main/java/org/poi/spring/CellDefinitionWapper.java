package org.poi.spring;

import org.poi.spring.config.ColumnDefinition;

/**
 * Created by Hong.LvHang on 2017-05-18.
 */
public class CellDefinitionWapper {
    //cell definition
    private ColumnDefinition columnDefinition;
    //cell interceptor errMessage
    private String errMessage;

    public CellDefinitionWapper(ColumnDefinition columnDefinition) {
        this.columnDefinition = columnDefinition;
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
