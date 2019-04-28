package com.alibaba.excel.metadata;

import org.apache.poi.ss.usermodel.CellStyle;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * @author jipengfei
 */
public class ExcelColumnProperty implements Comparable<ExcelColumnProperty> {

    /**
     */
    private Field field;

    /**
     */
    private int index = 99999;

    /**
     */
    private List<String> head = new ArrayList<String>();

    /**
     */
    private String format;

    /**
     * 类型转换器
     */
    private TypeConverter converter;

    /**
     * 单元格样式 错误场景样式
     */
    private Class<? extends CustomCellStyle> errorCellStyleClass;

    /**
     * 单元格样式 常规场景样式
     */
    private Class<? extends CustomCellStyle> commonCellStyleClass;

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public List<String> getHead() {
        return head;
    }

    public void setHead(List<String> head) {
        this.head = head;
    }

    public TypeConverter getConverter() {
        return converter;
    }

    public void setConverter(TypeConverter converter) {
        this.converter = converter;
    }

    public Class<? extends CustomCellStyle> getErrorCellStyleClass() {
        return errorCellStyleClass;
    }

    public void setErrorCellStyleClass(Class<? extends CustomCellStyle> errorCellStyleClass) {
        this.errorCellStyleClass = errorCellStyleClass;
    }

    public Class<? extends CustomCellStyle> getCommonCellStyleClass() {
        return commonCellStyleClass;
    }

    public void setCommonCellStyleClass(Class<? extends CustomCellStyle> commonCellStyleClass) {
        this.commonCellStyleClass = commonCellStyleClass;
    }

    public int compareTo(ExcelColumnProperty o) {
        int x = this.index;
        int y = o.getIndex();
        return (x < y) ? -1 : ((x == y) ? 0 : 1);
    }
}