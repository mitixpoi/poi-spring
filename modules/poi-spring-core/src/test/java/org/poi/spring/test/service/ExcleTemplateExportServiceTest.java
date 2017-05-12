package org.poi.spring.test.service;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.poi.spring.component.TemplateExcleHeader;
import org.poi.spring.service.ExcleTemplateExportService;
import org.poi.spring.service.impl.ExcleTemplateExportServiceImpl;
import org.poi.spring.service.result.ExcelExportResult;
import org.poi.spring.test.Car;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Hong.LvHang on 2017-05-12.
 */


/**
 * 测试一下
 * Created by Hong.LvHang on 2017-05-09.
 */
@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations = "classpath:applicationContext.xml")
public class ExcleTemplateExportServiceTest {

    private ExcleTemplateExportService excleTemplateExportService;

    /**
     * 测试模版导出
     */
    @Test
    public void createTemplateExceltest() {
        Car car1 = new Car("info1", "22", "VN222888IIN5", "卡卡西");
        ExcelExportResult result = ((ExcleTemplateExportServiceImpl) excleTemplateExportService).addExcelHeader(new TemplateExcleHeader() {
            @Override
            public String getHeader() {
                return "汽车销售表模版";
            }
        }).createTemplateExcel(car1);

        try {
            FileOutputStream fileOut = new FileOutputStream("d:/workbooktmp.xlsx");
            result.build().write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                result.getWorkbook().close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
