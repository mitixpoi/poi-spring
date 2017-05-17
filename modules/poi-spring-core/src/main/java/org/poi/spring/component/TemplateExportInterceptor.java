package org.poi.spring.component;

import org.poi.spring.config.ColumnDefinition;
import org.springframework.core.Ordered;

/**
 * Created by Hong.LvHang on 2017-05-17.
 */
public interface TemplateExportInterceptor extends Ordered {
    //导出前的处理
    String perHandle(ColumnDefinition columnDefinition, String value);

    //todo  可以做导出以后的处理  目前没有去实现
    //void postHandle();
}
