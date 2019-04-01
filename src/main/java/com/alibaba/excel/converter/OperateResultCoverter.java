package com.alibaba.excel.converter;

import com.alibaba.excel.metadata.RowErrorModel;
import com.alibaba.excel.metadata.TypeConverter;

import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * Description:
 *
 * @author cjb
 * @version V1.0
 * @since 2019-03-31 18:57
 */
public class OperateResultCoverter implements TypeConverter {

    @Override
    public Object convertToRead(String input) {
        // 读excel,不需要要转
        return null;
    }

    @Override
    public String convertToWrite(Object value) {
        if (value instanceof List) {
            List<Object> list = (List<Object>) value;
//            Integer lineNumber = (Integer) list.get(0);
            Map<String, RowErrorModel.CellInfo> errorMap = (Map<String, RowErrorModel.CellInfo>) list.get(1);
            StringBuilder stringBuilder = new StringBuilder();

            for (Map.Entry<String, RowErrorModel.CellInfo> entry : errorMap.entrySet()) {
                String title = entry.getKey();
                RowErrorModel.CellInfo cellInfo = entry.getValue();

                if (title.startsWith(RowErrorModel.INTERNAL_STRING)) {
                    stringBuilder.append(cellInfo.getRemark())
                            .append(";");
                    continue;
                }

                stringBuilder.append(title)
                        .append("=")
                        .append(cellInfo.getSource())
                        .append(",")
                        .append(cellInfo.getRemark())
                        .append(";\n\t");
            }
            return stringBuilder.toString();
        }

        return null;
    }
}

