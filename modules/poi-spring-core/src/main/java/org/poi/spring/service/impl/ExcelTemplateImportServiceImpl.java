package org.poi.spring.service.impl;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.poi.spring.CellReferenceWapper;
import org.poi.spring.ExcleWapper;
import org.poi.spring.ReflectUtil;
import org.poi.spring.component.ExcleContext;
import org.poi.spring.component.interceptor.TemplateImportInterceptor;
import org.poi.spring.config.ColumnDefinition;
import org.poi.spring.config.ExcelWorkBookBeandefinition;
import org.poi.spring.exception.ExcelException;
import org.poi.spring.service.ExcelTemplateImportService;
import org.poi.spring.service.result.ExcelImportResult;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.convert.TypeDescriptor;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Hong.LvHang on 2017-05-09.
 */
@Service
public class ExcelTemplateImportServiceImpl implements ExcelTemplateImportService {
    @Autowired
    private ExcleContext excleContext;

    /**
     * 格式化信息
     */
    private DataFormatter formatter = new DataFormatter();

    @Override
    public ExcelImportResult readExcel(Class<?> beanClass, InputStream excelStream) {
        if (!excleContext.exists(beanClass)) {
            throw new ExcelException("未找到匹配的模板");
        }
        //从注册信息中获取Bean信息
        ExcelWorkBookBeandefinition excelWorkBookBeandefinition = excleContext.getExcelWorkBookBeandefinition(beanClass);
        return doReadExcel(excelWorkBookBeandefinition, excelStream);
    }

    private ExcelImportResult doReadExcel(ExcelWorkBookBeandefinition excelWorkBookBeandefinition, InputStream excelStream) {
        ExcelImportResult result = null;
        Workbook workbook = null;

        try {
            workbook = WorkbookFactory.create(excelStream);

            Sheet sheet = workbook.getSheetAt(excelWorkBookBeandefinition.getSheetIndex());

            ExcleWapper wapper = new ExcleWapper(workbook, sheet);
            //标题之前的数据处理 目前没有用处  所以不处理
            //List<List<Object>> header = readHeader(excelWorkBookBeandefinition, sheet, titleIndex);
            //根据模版判断tital的index
            int titleIndex;
            if (excelWorkBookBeandefinition.getHeader() != null && excelWorkBookBeandefinition.getHeader() != "") {
                titleIndex = 1;
            } else {
                titleIndex = 0;
            }
            //获取标题
            Map<Integer, CellReferenceWapper> titleWapperMap = readTitle(excelWorkBookBeandefinition, wapper, titleIndex);
            //获取Bean
            List<Object> listBeans = readRows(excelWorkBookBeandefinition, wapper, titleIndex, titleWapperMap);

            result = new ExcelImportResult(listBeans);

        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return result;
    }

    private Map<Integer, CellReferenceWapper> readTitle(ExcelWorkBookBeandefinition excelWorkBookBeandefinition, ExcleWapper wapper, int titleIndex) {
        Map<Integer, CellReferenceWapper> map = new HashMap<>();
        List<ColumnDefinition> columnDefinitions = excelWorkBookBeandefinition.getColumnDefinitions();
        // 获取Excel标题数据
        Row row = wapper.getSheet().getRow(titleIndex);
        for (Cell cell : row) {
            //cell封装对象，可以获取cell的相关信息
            CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
            String text = formatter.formatCellValue(cell);
            for (ColumnDefinition columnDefinition : columnDefinitions) {
                if (columnDefinition.getTitle().equals(text)) {
                    map.put(cell.getColumnIndex(), new CellReferenceWapper(cellRef, columnDefinition));
                    break;
                }
            }
        }
        return map;
    }

    private <T> List<T> readRows(ExcelWorkBookBeandefinition excelWorkBookBeandefinition, ExcleWapper wapper, int titleIndex, Map<Integer, CellReferenceWapper> titleWapperMap) {
        int rowNum = wapper.getSheet().getLastRowNum();
        //读取数据的总共次数
        int totalNum = rowNum - titleIndex;
        List<T> listBean = new ArrayList<T>(totalNum);
        for (int i = titleIndex + 1; i <= rowNum; i++) {
            Object bean = readRow(excelWorkBookBeandefinition, wapper, i, titleWapperMap);
            listBean.add((T) bean);
        }
        return listBean;
    }

    private Object readRow(ExcelWorkBookBeandefinition excelWorkBookBeandefinition, ExcleWapper wapper, int rowNum, Map<Integer, CellReferenceWapper> titleWapperMap) {
        Row row = wapper.getSheet().getRow(rowNum);
        //创建注册时配置的bean类型
        Class<?> dataClass = excelWorkBookBeandefinition.getDataClass();
        Object bean = ReflectUtil.newInstance(dataClass);
        //默认的Cell迭代器回会过滤空格  按照传统方式进行循环
        int size = titleWapperMap.size();
        for (int index = 0; index < size; index++) {
            Cell cell = row.getCell(index, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            CellReferenceWapper titleWapper = titleWapperMap.get(cell.getColumnIndex());
            String text = formatter.formatCellValue(cell);
            //导入先拦截再转换
            doTemplateImportInterceptor(titleWapper, text);
            Object value = doConverter(bean, titleWapper, text);

            ReflectUtil.setProperty(bean, titleWapper.getColumnDefinition().getName(), value);
        }
        return bean;
    }

    private void doTemplateImportInterceptor(CellReferenceWapper titleWapper, String value) {
        String valueToUse = value;
        List<TemplateImportInterceptor> interceptors = excleContext.getTemplateImportInterceptors();
        if (interceptors != null && interceptors.size() > 0) {
            //执行操作
            for (TemplateImportInterceptor interceptor : interceptors) {
                if (!interceptor.perHandle(titleWapper, valueToUse)) {
                    throw new ExcelException(titleWapper.getErrMessage());
                }
            }
        }
    }

    private Object doConverter(Object bean, CellReferenceWapper titleWapper, String text) {
        if (text == null) {
            return null;
        }
        Class<?> fieldClass = ReflectUtil.getPropertyType(bean, titleWapper.getColumnDefinition().getName());
        if (!excleContext.getExcleConverter().canConvert(TypeDescriptor.valueOf(String.class), TypeDescriptor.valueOf(fieldClass))) {
            throw new ExcelException("字符串无法转换成对象值field=" + titleWapper.getColumnDefinition().getName());
        }
        return excleContext.getExcleConverter().convert(text, fieldClass);
    }

}
