package org.poi.spring.test;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.poi.spring.annotation.Column;
import org.poi.spring.annotation.Excle;

/**
 * Created by oldflame on 2017/4/8.
 */
@Excle(name = "导出文件1", width = 2, fgcolor = IndexedColors.YELLOW, border = BorderStyle.THIN, align = HorizontalAlignment.GENERAL, header = "1.单据不行需要整理 2.看的人是猪")
public class Car {

    @Column(title = "名字", width = 6, align = HorizontalAlignment.LEFT)
    public String name;

    @Column(title = "车龄")
    private String age;

    @Column(title = "车架号", align = HorizontalAlignment.RIGHT)
    private String vin;

    @Column(title = "客户")
    private String customerName;

    @Column(title = "客户状态", format = "0:正常,1:禁用", dictNo = "1010")
    private String customerStatus = "1";

    public Car() {

    }


    public Car(String name, String age, String vin, String customerName) {
        this.name = name;
        this.age = age;
        this.vin = vin;
        this.customerName = customerName;
    }

    public String getVin() {
        return vin;
    }

    public void setVin(String vin) {
        this.vin = vin;
    }

    public String getCustomerName() {
        return customerName;
    }

    public void setCustomerName(String customerName) {
        this.customerName = customerName;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getCustomerStatus() {
        return customerStatus;
    }

    public void setCustomerStatus(String customerStatus) {
        this.customerStatus = customerStatus;
    }
}
