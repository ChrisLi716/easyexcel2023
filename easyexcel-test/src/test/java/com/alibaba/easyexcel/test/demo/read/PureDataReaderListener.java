package com.alibaba.easyexcel.test.demo.read;

import cn.hutool.json.JSONUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Slf4j
public class PureDataReaderListener extends AnalysisEventListener<Map<Integer, String>> {

    private List<Map<Integer, Map<Integer, String>>> datas = new ArrayList<>();
    private Map<Integer, String> headTitleMap = new HashMap<>();

    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        headTitleMap = headMap;
    }

    @Override
    public void invoke(Map<Integer, String> data, AnalysisContext context) {
        Map<Integer, Map<Integer, String>> map = new HashMap<>();
        map.put(context.readRowHolder().getRowIndex(), data);
        datas.add(map);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        System.out.println("headTitleMap:" + JSONUtil.toJsonStr(headTitleMap));
        System.out.println("content:" + JSONUtil.toJsonStr(datas));
    }
}
