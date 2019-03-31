package com.alibaba.excel.metadata;

import com.alibaba.excel.annotation.ExcelProperty;

import java.util.HashMap;
import java.util.Map;

/**
 * Description:  行解析对应校验结果
 *
 * @author cjb
 * @version V1.0
 * @since 2019-03-30 15:42
 */
public class RowErrorModel extends BaseRowModel{

    // 行号
    private Integer lineNumber;

    // 错误信息Map
    private Map<String, CellInfo> errorMap = new HashMap<String, CellInfo>();

    // 操作结果
    @ExcelProperty(index = Integer.MAX_VALUE, value = {"===操作结果==="})
    private String operateResult;


    public Integer getLineNumber() {
        return lineNumber;
    }

    public void setLineNumber(Integer lineNumber) {
        this.lineNumber = lineNumber;
    }

    public Map<String, CellInfo> getErrorMap() {
        return errorMap;
    }

    public void setErrorMap(Map<String, CellInfo> errorMap) {
        this.errorMap = errorMap;
    }

    public String getOperateResult() {
        return operateResult;
    }

    public void setOperateResult(String operateResult) {
        this.operateResult = operateResult;
    }

    @Override
    public String toString() {
        final StringBuffer sb = new StringBuffer("RowErrorModel [");
        sb.append("lineNumber=").append(lineNumber);
        sb.append(", errorMap=").append(errorMap);
        sb.append(", operateResult=").append(operateResult);
        sb.append("]");
        return sb.toString();
    }

    public static class CellInfo {

        private Object source;

        private String remark;

        public Object getSource() {
            return source;
        }


        public String getRemark() {
            return remark;
        }

        public static CellInfo newInstance(Object source, String remark) {
            CellInfo cellInfo = new CellInfo();
            cellInfo.remark = remark;
            cellInfo.source = source;
            return cellInfo;
        }

        @Override
        public String toString() {
            final StringBuffer sb = new StringBuffer("CellInfo [");
            sb.append("source=").append(source);
            sb.append(", remark=").append(remark);
            sb.append("]");
            return sb.toString();
        }
    }
}

