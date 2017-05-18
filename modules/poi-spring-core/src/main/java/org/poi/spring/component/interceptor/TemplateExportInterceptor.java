package org.poi.spring.component.interceptor;

import org.poi.spring.CellDefinitionWapper;
import org.springframework.core.Ordered;

/**
 * Created by Hong.LvHang on 2017-05-17.
 */
public interface TemplateExportInterceptor extends Ordered {
    /**
     * 数据到处前处理
     * 1.处理返回到处的数据格式
     * 2.如果处理完成以后返回为空   那阻断停止导出
     * 3.处理错误放在errMessgae抛出
     *
     * @param cellDefinitionWapper
     * @param value
     * @return
     */
    String perHandle(CellDefinitionWapper cellDefinitionWapper, String value);

    //todo  可以做导出以后的处理  目前没有去实现
    //void postHandle();
}
