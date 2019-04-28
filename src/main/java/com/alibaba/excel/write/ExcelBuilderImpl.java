package com.alibaba.excel.write;

import com.alibaba.excel.context.WriteContext;
import com.alibaba.excel.event.WriteHandler;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.alibaba.excel.exception.ExcelGenerateException;
import com.alibaba.excel.metadata.*;
import com.alibaba.excel.style.LastColumnErrorCustomCellStyle;
import com.alibaba.excel.style.RowErrorCustomCellStyle;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.util.*;
import net.sf.cglib.beans.BeanMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Modifier;
import java.util.List;

/**
 * @author jipengfei
 */
public class ExcelBuilderImpl implements ExcelBuilder {

    private WriteContext context;

    public ExcelBuilderImpl(InputStream templateInputStream,
                            OutputStream out,
                            ExcelTypeEnum excelType,
                            boolean needHead, WriteHandler writeHandler) {
        try {
            //初始化时候创建临时缓存目录，用于规避POI在并发写bug
            POITempFile.createPOIFilesDirectory();
            context = new WriteContext(templateInputStream, out, excelType, needHead, writeHandler);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void addContent(List data, int startRow) {
        if (CollectionUtils.isEmpty(data)) {
            return;
        }
        int rowNum = context.getCurrentSheet().getLastRowNum();
        if (rowNum == 0) {
            Row row = context.getCurrentSheet().getRow(0);
            if (row == null) {
                if (context.getExcelHeadProperty() == null || !context.needHead()) {
                    rowNum = -1;
                }
            }
        }
        if (rowNum < startRow) {
            rowNum = startRow;
        }
        for (int i = 0; i < data.size(); i++) {
            int n = i + rowNum + 1;
            addOneRowOfDataToExcel(data.get(i), n);
        }
    }

    @Override
    public void addContent(List data, Sheet sheetParam) {
        context.currentSheet(sheetParam);
        addContent(data, sheetParam.getStartRow());
    }

    @Override
    public void addContent(List data, Sheet sheetParam, Table table) {
        context.currentSheet(sheetParam);
        context.currentTable(table);
        addContent(data, sheetParam.getStartRow());
    }

    @Override
    public void merge(int firstRow, int lastRow, int firstCol, int lastCol) {
        CellRangeAddress cra = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        context.getCurrentSheet().addMergedRegion(cra);
    }

    @Override
    public void finish() {
        try {
            context.getWorkbook().write(context.getOutputStream());
            context.getWorkbook().close();
        } catch (IOException e) {
            throw new ExcelGenerateException("IO error", e);
        }
    }

    private void addBasicTypeToExcel(List<Object> oneRowData, Row row) {
        if (CollectionUtils.isEmpty(oneRowData)) {
            return;
        }
        for (int i = 0; i < oneRowData.size(); i++) {
            Object cellValue = oneRowData.get(i);
            Cell cell = WorkBookUtil.createCell(row, i, context.getCurrentContentStyle(), cellValue,
                TypeUtil.isNum(cellValue));
            if (null != context.getAfterWriteHandler()) {
                context.getAfterWriteHandler().cell(i, cell);
            }
        }
    }

    private void addJavaObjectToExcel(Object oneRowData, Row row) {
        int i = 0;
        BeanMap beanMap = BeanMap.create(oneRowData);
        for (ExcelColumnProperty excelHeadProperty : context.getExcelHeadProperty().getColumnPropertyList()) {
            RowErrorModel baseRowModel = (RowErrorModel)oneRowData;
            String cellValue = TypeUtil.getFieldStringValue(beanMap, excelHeadProperty.getField().getName(),
                excelHeadProperty.getFormat(), excelHeadProperty.getConverter(), excelHeadProperty.getHead());
//            CellStyle cellStyle = baseRowModel.getStyle(i) != null ? baseRowModel.getStyle(i)
//                : context.getCurrentContentStyle();

            // 单元格标色，如果业务方没有给定对应的列名称，很难标识具体的单元格，因此这里改用行标识
            CustomCellStyle customCellStyle;
            if (baseRowModel.getErrorMap().size() == 0) {
                customCellStyle = getCellStyle(excelHeadProperty.getCommonCellStyleClass());
            } else {
                customCellStyle = getCellStyle(excelHeadProperty.getErrorCellStyleClass());
            }
            CellStyle currentCellStyle = null;
            if (customCellStyle != null) {
                currentCellStyle = customCellStyle.getStyle(context.getWorkbook());
            }

            // 错误场景下行样式
//            CustomCellStyle customCellStyle = null;
//            if (baseRowModel.getErrorMap().size() != 0) {
//                customCellStyle = new RowErrorCustomCellStyle();
//                if (excelHeadProperty.getField().getName().equalsIgnoreCase("operateResult")) {
//                    customCellStyle = new LastColumnErrorCustomCellStyle();
//                }
//            }

            Boolean num = TypeUtil.isNum(excelHeadProperty.getField());
            if (num && !StringUtils.isDigit(cellValue)) {
                // 格式不正确导致的Field 与 cellValue不匹配
                num = false;
            }
            Cell cell = WorkBookUtil.createCell(row, i, currentCellStyle, cellValue,
                    num);
            if (null != context.getAfterWriteHandler()) {
                context.getAfterWriteHandler().cell(i, cell);
            }
            i++;
        }

    }

    private CustomCellStyle getCellStyle(Class<? extends CustomCellStyle> customCellStyle) {
        try {
            if (!customCellStyle.equals(CustomCellStyle.class) && (customCellStyle.isInterface() || Modifier.isAbstract(customCellStyle.getModifiers()))) {
                throw new ExcelAnalysisException("customCellStyle type is wrong");
            }

            if (!customCellStyle.equals(CustomCellStyle.class)) {
                Class<? extends CustomCellStyle> custom = customCellStyle;
                return custom.newInstance();
            }
        } catch (InstantiationException e) {
            throw new ExcelAnalysisException("InstantiationException is happen : {}" , e);
        } catch (IllegalAccessException e) {
            throw new ExcelAnalysisException("IllegalAccessException is happen : {}" , e);
        }
        return null;
    }

    private void addOneRowOfDataToExcel(Object oneRowData, int n) {
        Row row = WorkBookUtil.createRow(context.getCurrentSheet(), n);
        if (null != context.getAfterWriteHandler()) {
            context.getAfterWriteHandler().row(n, row);
        }
        if (oneRowData instanceof List) {
            addBasicTypeToExcel((List)oneRowData, row);
        } else {
            addJavaObjectToExcel(oneRowData, row);
        }
    }
}
