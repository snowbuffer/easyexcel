package com.alibaba.excel.style;

import com.alibaba.excel.metadata.CustomCellStyle;
import org.apache.poi.ss.usermodel.*;

/**
 * Description: 单元格错误标色
 *
 * @author cjb
 * @version V1.0
 * @since 2019-04-28 11:46
 */
public class LastColumnErrorCustomCellStyle implements CustomCellStyle {

    @Override
    public CellStyle getStyle(Workbook workbook) {
        CellStyle newCellStyle = getBaseStyle(workbook);
        if (newCellStyle == null) {
            newCellStyle = workbook.createCellStyle();
        }
        // 自定义样式
        Font font = workbook.createFont();
        font.setColor(Font.COLOR_RED);
        newCellStyle.setFont(font);
        return newCellStyle;
    }

    @Override
    public CellStyle getBaseStyle(Workbook workbook) {
        // 基础样式
        CellStyle newCellStyle = workbook.createCellStyle();
        return newCellStyle;
    }

}

