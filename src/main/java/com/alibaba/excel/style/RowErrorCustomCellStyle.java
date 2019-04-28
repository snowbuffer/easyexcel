package com.alibaba.excel.style;

import com.alibaba.excel.metadata.CustomCellStyle;
import org.apache.poi.ss.usermodel.*;

/**
 * Description:
 *
 * @author cjb
 * @version V1.0
 * @since 2019-04-28 13:42
 */
public class RowErrorCustomCellStyle implements CustomCellStyle {

    @Override
    public CellStyle getStyle(Workbook workbook) {
        CellStyle newCellStyle = getBaseStyle(workbook);
        if (newCellStyle == null) {
            newCellStyle = workbook.createCellStyle();
        }
        // 自定义样式
        newCellStyle.setBorderBottom(BorderStyle.THIN);
        newCellStyle.setBorderLeft(BorderStyle.THIN);
        newCellStyle.setBorderRight(BorderStyle.THIN);
        newCellStyle.setBorderTop(BorderStyle.THIN);
        return newCellStyle;
    }

    @Override
    public CellStyle getBaseStyle(Workbook workbook) {
        // 基础样式
        CellStyle newCellStyle = workbook.createCellStyle();
        newCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return newCellStyle;
    }
}

