package com.alibaba.easyexcel.test;

import com.alibaba.easyexcel.test.listen.AfterWriteHandlerImpl;
import com.alibaba.easyexcel.test.model.WriteModel;
import com.alibaba.easyexcel.test.util.FileUtil;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.junit.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static com.alibaba.easyexcel.test.util.DataUtil.*;

public class WriteTest {

    @Test
    public void writeNew() throws IOException {
        OutputStream out = new FileOutputStream("E:/2007_new6.xlsx");
        ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX, true);
        for (int j = 1; j <= 50; j++) {
            com.alibaba.excel.metadata.Sheet sheet = new com.alibaba.excel.metadata.Sheet(j, 1);

            PointManager manager = new PointManager();
            manager.addPoint(new Point(3, 7, 1, 4, 5));
            manager.addPoint(new Point(2, 1, 2, 1, "2"));
            manager.addPoint(new Point(2, 3, 2, 3, "3"));
            manager.addPoint(new Point(2, 5, 3, 1, "4"));
            manager.addPoint(new Point(11, 5, 3, 1, "7"));
            manager.addPoint(new Point(1, 1, 7, 1, "1"));
            manager.addPoint(new Point(12, 2, 4, 2, "8"));
            manager.addPoint(new Point(6, 2, 4, 3, "6"));
            manager.calc();
            List<List<Object>> titleList = manager.getTitleList();
            sheet.setTitleList(titleList);

            List<List<String>> headList = new ArrayList<List<String>>();
            for (int i = 0; i < 10; i++) {
                List<String> temp =new ArrayList<String>();
                temp.add("title_" + i);
                headList.add(temp);
            }
            sheet.setSheetName("sheet_" + j);
            sheet.setHead(headList);
            sheet.setStartRow(0);

            List<List<Object>> allDataList = new ArrayList<List<Object>>();
            for (int m = 1; m <= 200; m++) {
                List<Object> tempList=  new ArrayList<Object>();
                for (int i = 0; i < 10; i++) {
                    tempList.add(sheet.getSheetName() + "_r_" + m + "_data_" + i);
                }
                allDataList.add(tempList);
            }
            writer.write1(allDataList, sheet);

            manager.getPointList().forEach(point -> {
                List<Integer> pointRangeList = point.getPointRangeList();

                writer.merge(
                        pointRangeList.get(0),
                        pointRangeList.get(1),
                        pointRangeList.get(2),
                        pointRangeList.get(3));
            });
            Map<Integer,Integer> columnWidthMap = new HashMap<>();
            columnWidthMap.put(0, 10);
            columnWidthMap.put(1, 20);
            columnWidthMap.put(2, 10);
            columnWidthMap.put(3, 20);
            columnWidthMap.put(4, 10);
            columnWidthMap.put(5, 20);
            columnWidthMap.put(6, 10);
            columnWidthMap.put(7, 20);
            columnWidthMap.put(8, 10);
            columnWidthMap.put(9, 20);
            writer.setColumnWidth(columnWidthMap, 10);
        }
        writer.finish();
        out.close();
    }


    @Test
    public void writeV2007() throws IOException {
        OutputStream out = new FileOutputStream("D:/2007.xlsx");
        ExcelWriter writer = EasyExcelFactory.getWriter(out);
        //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
        Sheet sheet1 = new Sheet(1, 3);
        sheet1.setSheetName("第一个sheet");

        //设置列宽 设置每列的宽度
        Map columnWidth = new HashMap();
        columnWidth.put(0,10000);columnWidth.put(1,40000);columnWidth.put(2,10000);columnWidth.put(3,10000);
        sheet1.setColumnWidthMap(columnWidth);
        sheet1.setHead(createTestListStringHead());
        //or 设置自适应宽度
        //sheet1.setAutoWidth(Boolean.TRUE);
        writer.write1(createTestListObject(), sheet1);

        //写第二个sheet sheet2  模型上打有表头的注解，合并单元格
        Sheet sheet2 = new Sheet(2, 3, WriteModel.class, "第二个sheet", null);
        sheet2.setTableStyle(createTableStyle());
        //writer.write1(null, sheet2);
        writer.write(createTestListJavaMode(), sheet2);
        //需要合并单元格
        writer.merge(5,20,1,1);

        //写第三个sheet包含多个table情况
        Sheet sheet3 = new Sheet(3, 0);
        sheet3.setSheetName("第三个sheet");
        Table table1 = new Table(1);
        table1.setHead(createTestListStringHead());
        writer.write1(createTestListObject(), sheet3, table1);

        //写sheet2  模型上打有表头的注解
        Table table2 = new Table(2);
        table2.setTableStyle(createTableStyle());
        table2.setClazz(WriteModel.class);
        writer.write(createTestListJavaMode(), sheet3, table2);

        writer.finish();
        out.close();

    }


    @Test
    public void writeV2007WithTemplate() throws IOException {
        InputStream inputStream = FileUtil.getResourcesFileInputStream("temp.xlsx");
        OutputStream out = new FileOutputStream("D:/2007.xlsx");
        ExcelWriter writer = EasyExcelFactory.getWriterWithTemp(inputStream,out,ExcelTypeEnum.XLSX,true);
//        //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
//        Sheet sheet1 = new Sheet(1, 3);
//        sheet1.setSheetName("第一个sheet");
//        sheet1.setStartRow(20);
//
//        //设置列宽 设置每列的宽度
//        Map columnWidth = new HashMap();
//        columnWidth.put(0,10000);columnWidth.put(1,40000);columnWidth.put(2,10000);columnWidth.put(3,10000);
//        sheet1.setColumnWidthMap(columnWidth);
//        sheet1.setHead(createTestListStringHead());
//        //or 设置自适应宽度
//        //sheet1.setAutoWidth(Boolean.TRUE);
//        writer.write1(createTestListObject(), sheet1);
//
//        //写第二个sheet sheet2  模型上打有表头的注解，合并单元格
//        Sheet sheet2 = new Sheet(2, 3, WriteModel.class, "第二个sheet", null);
//        sheet2.setTableStyle(createTableStyle());
//        sheet2.setStartRow(20);
//        writer.write(createTestListJavaMode(), sheet2);
//
        //写第三个sheet包含多个table情况
        Sheet sheet3 = new Sheet(3, 0);
        sheet3.setSheetName("第三个sheet");
        sheet3.setStartRow(30);
        Table table1 = new Table(1);
        table1.setHead(createTestListStringHead());
        writer.write1(createTestListObject(), sheet3, table1);

        //写sheet2  模型上打有表头的注解
        Table table2 = new Table(2);
        table2.setTableStyle(createTableStyle());
        table2.setClazz(WriteModel.class);
        writer.write(createTestListJavaMode(), sheet3, table2);

        writer.finish();
        out.close();

    }

    @Test
    public void writeV2007WithTemplateAndHandler() throws IOException {
        InputStream inputStream = FileUtil.getResourcesFileInputStream("temp.xlsx");
        OutputStream out = new FileOutputStream("/Users/jipengfei/2007.xlsx");
        ExcelWriter writer = EasyExcelFactory.getWriterWithTempAndHandler(inputStream,out,ExcelTypeEnum.XLSX,true,
            new AfterWriteHandlerImpl());
        //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
        Sheet sheet1 = new Sheet(1, 3);
        sheet1.setSheetName("第一个sheet");
        sheet1.setStartRow(20);

        //设置列宽 设置每列的宽度
        Map columnWidth = new HashMap();
        columnWidth.put(0,10000);columnWidth.put(1,40000);columnWidth.put(2,10000);columnWidth.put(3,10000);
        sheet1.setColumnWidthMap(columnWidth);
        sheet1.setHead(createTestListStringHead());
        //or 设置自适应宽度
        //sheet1.setAutoWidth(Boolean.TRUE);
        writer.write1(createTestListObject(), sheet1);

        //写第二个sheet sheet2  模型上打有表头的注解，合并单元格
        Sheet sheet2 = new Sheet(2, 3, WriteModel.class, "第二个sheet", null);
        sheet2.setTableStyle(createTableStyle());
        sheet2.setStartRow(20);
        writer.write(createTestListJavaMode(), sheet2);

        //写第三个sheet包含多个table情况
        Sheet sheet3 = new Sheet(3, 0);
        sheet3.setSheetName("第三个sheet");
        sheet3.setStartRow(30);
        Table table1 = new Table(1);
        table1.setHead(createTestListStringHead());
        writer.write1(createTestListObject(), sheet3, table1);

        //写sheet2  模型上打有表头的注解
        Table table2 = new Table(2);
        table2.setTableStyle(createTableStyle());
        table2.setClazz(WriteModel.class);
        writer.write(createTestListJavaMode(), sheet3, table2);

        writer.finish();
        out.close();

    }



    @Test
    public void writeV2003() throws IOException {
        OutputStream out = new FileOutputStream("/Users/jipengfei/2003.xls");
        ExcelWriter writer = EasyExcelFactory.getWriter(out, ExcelTypeEnum.XLS,true);
        //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
        Sheet sheet1 = new Sheet(1, 3);
        sheet1.setSheetName("第一个sheet");

        //设置列宽 设置每列的宽度
        Map columnWidth = new HashMap();
        columnWidth.put(0,10000);columnWidth.put(1,40000);columnWidth.put(2,10000);columnWidth.put(3,10000);
        sheet1.setColumnWidthMap(columnWidth);
        sheet1.setHead(createTestListStringHead());
        //or 设置自适应宽度
        //sheet1.setAutoWidth(Boolean.TRUE);
        writer.write1(createTestListObject(), sheet1);

        //写第二个sheet sheet2  模型上打有表头的注解，合并单元格
        Sheet sheet2 = new Sheet(2, 3, WriteModel.class, "第二个sheet", null);
        sheet2.setTableStyle(createTableStyle());
        writer.write(createTestListJavaMode(), sheet2);

        //写第三个sheet包含多个table情况
        Sheet sheet3 = new Sheet(3, 0);
        sheet3.setSheetName("第三个sheet");
        Table table1 = new Table(1);
        table1.setHead(createTestListStringHead());
        writer.write1(createTestListObject(), sheet3, table1);

        //写sheet2  模型上打有表头的注解
        Table table2 = new Table(2);
        table2.setTableStyle(createTableStyle());
        table2.setClazz(WriteModel.class);
        writer.write(createTestListJavaMode(), sheet3, table2);

        writer.finish();
        out.close();
    }
}
