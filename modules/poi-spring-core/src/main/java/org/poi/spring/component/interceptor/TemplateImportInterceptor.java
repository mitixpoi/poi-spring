package org.poi.spring.component.interceptor;

import org.poi.spring.CellReferenceWapper;
import org.springframework.core.Ordered;

/**
 * Created by Hong.LvHang on 2017-05-17.
 */
public interface TemplateImportInterceptor extends Ordered {
    /**
     * 数据从Excle倒入前的前置处理
     * 1.前置处理返回false则前置处理失败  倒入中断
     * 2.倒入处理失败返回的信息放入CellReferenceWapper  外面进行处理
     *
     * @param cellReferenceWapper
     * @param value
     * @return
     */
    boolean perHandle(CellReferenceWapper cellReferenceWapper, String value);
}
