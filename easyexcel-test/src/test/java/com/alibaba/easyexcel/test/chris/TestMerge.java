package com.alibaba.easyexcel.test.chris;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.easyexcel.test.demo.fill.FillData;
import com.alibaba.easyexcel.test.demo.fill.FillTest;
import com.alibaba.easyexcel.test.util.TestFileUtil;
import com.alibaba.excel.EasyExcel;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.Date;
import java.util.List;

@Slf4j
public class TestMerge {

    @Test
    public void testMergeCell() {
        String templateFileName =
            TestFileUtil.getPath() + "demo" + File.separator + "fill" + File.separator + "list.xlsx";

        List<FillData> data = data();
        // 方案1 一下子全部放到内存里面 并填充
        String fileName = TestFileUtil.getPath() + "listFill" + System.currentTimeMillis() + ".xlsx";
        // 这里 会填充到第一个sheet， 然后文件流会自动关闭
        EasyExcel.write(fileName).withTemplate(templateFileName).registerWriteHandler(new MergeCellStrategy()).sheet().doFill(data);
    }


    @Test
    public void testMergeRow() {
        String templateFileName =
            TestFileUtil.getPath() + "demo" + File.separator + "fill" + File.separator + "list.xlsx";

        List<FillData> data = data();
        String fileName = TestFileUtil.getPath() + "listFill" + System.currentTimeMillis() + ".xlsx";

        // 这里 会填充到第一个sheet， 然后文件流会自动关闭
        EasyExcel.write(fileName)
            .withTemplate(templateFileName)
            .registerWriteHandler(new MergeCellHandler().setMergeColumnIndex(new int[]{0, 1, 2}))
            .sheet()
            .doFill(data);
    }


    public static List<FillData> data() {
        List<FillData> list = CollUtil.newArrayList();

        FillData data01 = new FillData();
        data01.setName("张三");
        data01.setNumber(1);
        data01.setDate(new Date());

        FillData data02 = new FillData();
        data02.setName("李四");
        data02.setNumber(2);
        data02.setDate(new Date());

        FillData data03 = new FillData();
        data03.setName("张三");
        data03.setNumber(2);
        data03.setDate(new Date());

        list.add(data01);
        list.add(data03);
        list.add(data02);

        return list;
    }
}
