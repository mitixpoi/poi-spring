package org.poi.spring.service.impl;

import org.poi.spring.component.ExcelDictService;
import org.springframework.stereotype.Service;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by jianxunxie on 2017/5/12.
 */
@Service
public class ExcelDictServiceImpl implements ExcelDictService {

    @Override
    public Map<String, String> getColumnDictNoMap(String dictNo) {
        Map<String, String> map = new HashMap<>();
        map.put("1", "JAVA");
        map.put("2", "C");
        return map;
    }
}
