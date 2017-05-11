package org.poi.spring.test;

import org.poi.spring.annotation.Align;
import org.poi.spring.annotation.Column;
import org.poi.spring.annotation.Excle;

/**
 * Created by oldflame on 2017/4/8.
 */
@Excle(name = "导出文件1",width = 2)
public class Car {

    @Column(title = "名字",width = 6)
    public String name;

    @Column(title = "车龄")
    private String age;

    public Car(){

    }


    public Car(String name,String age) {
        this.name = name;
        this.age = age;
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

}
