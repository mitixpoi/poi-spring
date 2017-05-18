package org.poi.spring.test.service;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.poi.spring.service.ExcelTemplateImportService;
import org.poi.spring.service.result.ExcelImportResult;
import org.poi.spring.test.Car;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;

/**
 * Created by Hong.LvHang on 2017-05-10.
 */
@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations = "classpath:applicationContext.xml")
public class ExcelTemplateImportServiceTest {

    @Autowired
    private ExcelTemplateImportService excelTemplateImportService;

    @Test
    public void readExceltest() throws FileNotFoundException {
        ExcelImportResult excelImportResult = excelTemplateImportService.readExcel(Car.class, new FileInputStream("d:/workbooktmp.xlsx"));
        List<Car> cars = (List<Car>) excelImportResult.getBeans();
        System.out.println(cars.get(0).getName());
    }
}
