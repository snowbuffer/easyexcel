package com.alibaba.excel.modelbuild;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelGenerateException;
import com.alibaba.excel.metadata.ExcelHeadProperty;
import com.alibaba.excel.metadata.RowErrorModel;
import com.alibaba.excel.util.TypeUtil;
import net.sf.cglib.beans.BeanMap;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author jipengfei
 */
public class ModelBuildEventListener extends AnalysisEventListener {

    @Override
    public void invoke(Object object, AnalysisContext context) {
        if (context.getExcelHeadProperty() != null && context.getExcelHeadProperty().getHeadClazz() != null) {
            try {
                Object resultModel = buildUserModel(context, (List<String>)object);
                context.setCurrentRowAnalysisResult(resultModel);
            } catch (Exception e) {
                throw new ExcelGenerateException(e);
            }
        }
    }

    private Object buildUserModel(AnalysisContext context, List<String> stringList) throws Exception {
        ExcelHeadProperty excelHeadProperty = context.getExcelHeadProperty();
        Object resultModel = excelHeadProperty.getHeadClazz().newInstance();
        if (excelHeadProperty == null) {
            return resultModel;
        }

        Map<String, RowErrorModel.CellInfo> rowErrorMap = new HashMap();
        BeanMap.create(resultModel).putAll(
            TypeUtil.getFieldValues(stringList, excelHeadProperty, context.use1904WindowDate(), rowErrorMap));

        if (resultModel instanceof RowErrorModel) {
            RowErrorModel rowErrorModel = (RowErrorModel) resultModel;
            rowErrorModel.setLineNumber(context.getCurrentRowNum());
            rowErrorModel.setErrorMap(rowErrorMap);
        }

        if (rowErrorMap.size() > 0) {
            context.setParseSuccess(false);
        }
        return resultModel;
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

    }
}
