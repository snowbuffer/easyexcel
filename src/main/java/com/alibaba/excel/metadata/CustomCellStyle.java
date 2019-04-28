package com.alibaba.excel.metadata;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Description:
 *
 * @author cjb
 * @version V1.0
 * @since 2019-04-28 11:07
 */
public interface CustomCellStyle {

    CellStyle getStyle(Workbook workbook);

    CellStyle getBaseStyle(Workbook workbook);

}
