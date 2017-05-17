package org.poi.spring.service.impl;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.poi.spring.ExcleWapper;
import org.poi.spring.PoiConstant;
import org.poi.spring.ReflectUtil;
import org.poi.spring.component.ExcelHeader;
import org.poi.spring.component.ExcleContext;
import org.poi.spring.component.TemplateExcleHeader;
import org.poi.spring.component.TemplateExportInterceptor;
import org.poi.spring.config.ColumnDefinition;
import org.poi.spring.config.ExcelWorkBookBeandefinition;
import org.poi.spring.exception.ExcelException;
import org.poi.spring.service.ExcleTemplateExportService;
import org.poi.spring.service.result.ExcelExportResult;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.OrderComparator;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Hong.LvHang on 2017-05-12.
 */
@Service
public class ExcleTemplateExportServiceImpl implements ExcleTemplateExportService {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelExportServiceImpl.class);

    @Autowired
    private ExcleContext excleContext;

    private ExcelHeader excelHeader;

    private static final Object EMPTY_OBJECT = null;

    @Override
    public ExcelExportResult createTemplateExcel(Object bean) {
        ExcelExportResult exportResult = null;
        if (bean != null) {
            Class<?> beanClass = bean.getClass();
            ExcelWorkBookBeandefinition excelWorkBookBeandefinition = getExcelWorkBookBeandefinition(beanClass);
            List<Object> beans = new ArrayList<>();
            beans.add(bean);
            //创建表格
            exportResult = doCreateExcel(excelWorkBookBeandefinition, beans);
        }
        return exportResult;
    }

    private ExcelExportResult doCreateExcel(ExcelWorkBookBeandefinition excelWorkBookBeandefinition, List<?> beans) {
        Workbook workbook = new XSSFWorkbook();
        //直接使用sheetname  在初始化的时候已经设置了名字
        Sheet sheet = workbook.createSheet(excelWorkBookBeandefinition.getSheetName());
        ExcleWapper wapper = new ExcleWapper(workbook, sheet);

        //初始化文件头
        doCreateHeader(wapper, excelWorkBookBeandefinition, beans);
        //设置title
        createTitle(excelWorkBookBeandefinition, wapper);

        //如果listBean不为空,创建数据行
        if (beans != null) {
            createRows(excelWorkBookBeandefinition, wapper, beans);
        }
        ExcelExportResult exportResult = new ExcelExportResult(excelWorkBookBeandefinition, sheet, workbook);
        return exportResult;
    }

    /**
     * 定制表头
     * 极简模式和自定义模式
     *
     * @param wapper
     * @param excelWorkBookBeandefinition
     * @param beans
     */
    //先调用addExcelHeader 可以注册表头生成方式
    private void doCreateHeader(ExcleWapper wapper, ExcelWorkBookBeandefinition excelWorkBookBeandefinition, List<?> beans) {
        if (excelHeader != null && excelHeader instanceof TemplateExcleHeader) {
            Row row = wapper.getSheet().createRow((short) 0);
            //设置表头的值
            Cell cell = row.createCell((short) 0);
            cell.setCellValue(((TemplateExcleHeader) excelHeader).getHeader());
            //合并单元格
            int maxSize = excelWorkBookBeandefinition.getColumnDefinitions().size() - 1;
            wapper.getSheet().addMergedRegion(new CellRangeAddress(0, 0, 0, maxSize >= 0 ? maxSize : 0));
            //获取样式
            Map<String, Object> properties = createTemplateHeader();
            //设置第一个单元格样式
            CellUtil.setCellStyleProperties(cell, properties);
            for (int i = 1; i <= maxSize; i++) {
                //设置后续单元格样式
                Cell cellMerge = row.createCell(i);
                CellUtil.setCellStyleProperties(cellMerge, properties);
            }
        }
    }

    /**
     * 设置极简表头样式
     *
     * @return
     */
    private Map<String, Object> createTemplateHeader() {
        Map<String, Object> properties = new HashMap<>();
        //设置底色
        properties.put(CellUtil.FILL_FOREGROUND_COLOR, IndexedColors.AQUA.getIndex());
        properties.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
        //设置边线样式
        properties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);
        return properties;
    }

    /**
     * excle标题
     *
     * @param excelWorkBookBeandefinition
     * @param wapper
     * @return
     */
    private void createTitle(ExcelWorkBookBeandefinition excelWorkBookBeandefinition, ExcleWapper wapper) {
        //标题索引号
        int titleIndex = wapper.getSheet().getPhysicalNumberOfRows();
        Row titleRow = wapper.getSheet().createRow(titleIndex);
        List<ColumnDefinition> columnDefinitions = excelWorkBookBeandefinition.getColumnDefinitions();
        for (int i = 0; i < columnDefinitions.size(); i++) {
            ColumnDefinition columnDefinition = columnDefinitions.get(i);
            //设置宽度 都是在表头中处理的
            if (columnDefinition.getColumnWidth() != null) {
                wapper.getSheet().setColumnWidth(i, columnDefinition.getColumnWidth());
            }
            Cell cell = titleRow.createCell(i);
            CellUtil.setCellStyleProperties(cell, columnDefinition.getProperties());
            cell.setCellValue(columnDefinition.getTitle());
        }
    }

    public void createRows(ExcelWorkBookBeandefinition excelWorkBookBeandefinition, ExcleWapper wapper, List<?> beans) {
        int startRow = wapper.getSheet().getPhysicalNumberOfRows();
        for (int i = 0; i < beans.size(); i++) {
            int rowNum = startRow + i;
            createRow(excelWorkBookBeandefinition, wapper, rowNum, beans.get(i));
        }
    }

    private void createRow(ExcelWorkBookBeandefinition excelWorkBookBeandefinition, ExcleWapper wapper, int rowNum, Object bean) {
        Row row = wapper.getSheet().createRow(rowNum);
        List<ColumnDefinition> columnDefinitions = excelWorkBookBeandefinition.getColumnDefinitions();
        for (int i = 0; i < columnDefinitions.size(); i++) {
            ColumnDefinition columnDefinition = columnDefinitions.get(i);
            String name = columnDefinition.getName();
            Object value = ReflectUtil.getProperty(bean, name);
            //数据转换
            String valueStr = doConverter(name, value);
            if (valueStr != null) {
                //数据拦截处理
                valueStr = doTemplateExportInterceptor(columnDefinition, valueStr);
                valueStr = doCreateValidationIfNecessary(columnDefinition, row, rowNum, valueStr);
            }
            Cell cell = row.createCell(i);
            cell.setCellValue(valueStr);
        }
    }

    private String doCreateValidationIfNecessary(ColumnDefinition columnDefinition, Row row, int rowNum, String valueStr) {
        if (!PoiConstant.EMPTY_STRING.equals(columnDefinition.getDictNo())
            || !PoiConstant.EMPTY_STRING.equals(columnDefinition.getFormat())) {
            String[] dictArr = new String[0];
            if (!PoiConstant.EMPTY_STRING.equals(columnDefinition.getDictNo())) {
                if (null != excleContext.getExcelDictService()) {
                    Map<String, String> map = excleContext.getExcelDictService().getColumnDictNoMap(columnDefinition.getDictNo());
                    valueStr = map.get(valueStr);
                    dictArr = new String[map.size()];
                    int index = 0;
                    for (String s : map.keySet()) {
                        dictArr[index++] = map.get(s);
                    }
                }
            } else if (!PoiConstant.EMPTY_STRING.equals(columnDefinition.getFormat())) {
                try {
                    String[] expressions = StringUtils.split(columnDefinition.getFormat(), ",");
                    dictArr = new String[expressions.length];
                    for (int j = 0; j < expressions.length; j++) {
                        String expression = expressions[j];
                        String[] val = StringUtils.split(expression, ":");
                        String v1 = val[0];
                        String v2 = val[1];
                        if (valueStr.equals(v1)) {
                            valueStr = v2;
                        }
                        dictArr[j] = v2;
                    }
                } catch (Exception e) {
                    throw new ExcelException("表达式:" + columnDefinition.getFormat() + "错误,正确的格式应该以[,]号分割,[:]号取值");
                }
            }
            XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) row.getSheet());
            XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper.createExplicitListConstraint(dictArr);
            CellRangeAddressList addressList =
                new CellRangeAddressList(rowNum, rowNum, row.getPhysicalNumberOfCells(), row.getPhysicalNumberOfCells());
            XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint, addressList);
            row.getSheet().addValidationData(validation);
        }
        return valueStr;
    }

    /**
     * 导出前的处理
     *
     * @param columnDefinition
     * @param valueStr
     * @return
     */
    private String doTemplateExportInterceptor(ColumnDefinition columnDefinition, String valueStr) {
        List<Object> interceptors = columnDefinition.getTemplateExportInterceptors();
        String valueStrToUse = valueStr;
        if (interceptors == null || interceptors.size() == 0) {
            return valueStrToUse;
        }
        List<TemplateExportInterceptor> interceptorsToUse = new ArrayList<>();
        for (Object interceptor : interceptors) {
            if (interceptor instanceof TemplateExportInterceptor) {
                interceptorsToUse.add((TemplateExportInterceptor) interceptor);
            } else if (interceptor instanceof String) {
                TemplateExportInterceptor interceptorEntity = excleContext.getTemplateExportInterceptorMap().get(interceptor);
                if (interceptorEntity != null) {
                    interceptorsToUse.add((TemplateExportInterceptor) interceptor);
                }
            } else {
                throw new ExcelException("错误的数据导出拦截处理器" + interceptor.toString());
            }
        }
        //排序
        Collections.sort(interceptorsToUse, OrderComparator.INSTANCE);
        for (TemplateExportInterceptor interceptor : interceptorsToUse) {
            valueStrToUse = interceptor.perHandle(columnDefinition, valueStrToUse);
            if (valueStrToUse == null) {
                throw new ExcelException("拦截器执行错误" + interceptor.toString());
            }
        }
        return valueStrToUse;
    }

    private String doConverter(String name, Object value) {
        String valueStr = null;
        if (EMPTY_OBJECT == null) {
            return valueStr;
        }
        if (!excleContext.getExcleConverter().canConvertString(value.getClass())) {
            throw new ExcelException("无法转换成字符串,字段名name" + name);
        }
        valueStr = excleContext.getExcleConverter().convert(value, String.class);
        return valueStr;
    }

    private ExcelWorkBookBeandefinition getExcelWorkBookBeandefinition(Class<?> beanClass) {
        if (!excleContext.exists(beanClass)) {
            throw new ExcelException("未找到匹配的模板--" + beanClass.getName());
        }
        //从注册信息中获取Bean信息
        return excleContext.getExcelWorkBookBeandefinition(beanClass);
    }

    public ExcleTemplateExportServiceImpl addExcelHeader(ExcelHeader excelHeader) {
        this.excelHeader = excelHeader;
        return this;
    }
}
