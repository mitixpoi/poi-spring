package org.poi.spring.component.interceptor;

import org.poi.spring.CellReferenceWapper;
import org.springframework.core.Ordered;
import org.springframework.util.StringUtils;

/**
 * Created by Hong.LvHang on 2017-05-18.
 */
public class RequiredImportInterceptor implements TemplateImportInterceptor {

    @Override
    public boolean perHandle(CellReferenceWapper cellReferenceWapper, String value) {
        if (cellReferenceWapper.getColumnDefinition().isRequired() && !StringUtils.hasText(value)) {
            cellReferenceWapper.setErrMessage("必填项" + cellReferenceWapper.getColumnDefinition().getTitle());
            return false;
        }
        return true;
    }

    @Override
    public int getOrder() {
        return Ordered.HIGHEST_PRECEDENCE;
    }


}
