package com.alibaba.excel.util;

import com.alibaba.excel.exception.ExcelAnalysisException;
import com.alibaba.excel.metadata.ExcelColumnProperty;
import com.alibaba.excel.metadata.ExcelHeadProperty;
import com.alibaba.excel.metadata.RowErrorModel;
import com.alibaba.excel.metadata.TypeConverter;
import net.sf.cglib.beans.BeanMap;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author jipengfei
 */
public class TypeUtil {

    private static List<String> DATE_FORMAT_LIST = new ArrayList<String>(4);

    static {
        DATE_FORMAT_LIST.add("yyyy/MM/dd HH:mm:ss");
        DATE_FORMAT_LIST.add("yyyy-MM-dd HH:mm:ss");
        DATE_FORMAT_LIST.add("yyyyMMdd HH:mm:ss");
    }

    private static int getCountOfChar(String value, char c) {
        int n = 0;
        if (value == null) {
            return 0;
        }
        char[] chars = value.toCharArray();
        for (char cc : chars) {
            if (cc == c) {
                n++;
            }
        }
        return n;
    }

    public static Object convert(String value, Field field, String format, TypeConverter converter, boolean us) {

        // 转换器存在情况下，直接走转换器
        if (converter != null) {
            return converter.convertToRead(value);
        }

        if (!StringUtils.isEmpty(value)) {
            if (Float.class.equals(field.getType())) {
                return Float.parseFloat(value);
            }
            if (Integer.class.equals(field.getType()) || int.class.equals(field.getType())) {
                return Integer.parseInt(value);
            }
            if (Double.class.equals(field.getType()) || double.class.equals(field.getType())) {
                if (null != format && !"".equals(format)) {
                    int n = getCountOfChar(value, '0');
                    return Double.parseDouble(TypeUtil.formatFloat0(value, n));
                } else {
                    return Double.parseDouble(TypeUtil.formatFloat(value));
                }
            }
            if (Boolean.class.equals(field.getType()) || boolean.class.equals(field.getType())) {
                String valueLower = value.toLowerCase();
                if (valueLower.equals("true") || valueLower.equals("false")) {
                    return Boolean.parseBoolean(value.toLowerCase());
                }
                Integer integer = Integer.parseInt(value);
                if (integer == 0) {
                    return false;
                } else {
                    return true;
                }
            }
            if (Long.class.equals(field.getType()) || long.class.equals(field.getType())) {
                return Long.parseLong(value);
            }
            if (Date.class.equals(field.getType())) {
                if (value.contains("-") || value.contains("/") || value.contains(":")) {
                    return getSimpleDateFormatDate(value, format);
                } else {
                    Double d = Double.parseDouble(value);
                    return HSSFDateUtil.getJavaDate(d, us);
                }
            }
            if (BigDecimal.class.equals(field.getType())) {
                return new BigDecimal(value);
            }
            if(String.class.equals(field.getType())){
                return formatFloat(value);
            }

        }
        return null;
    }

    public static Boolean isNum(Field field) {
        if (field == null) {
            return false;
        }
        if (Integer.class.equals(field.getType()) || int.class.equals(field.getType())) {
            return true;
        }
        if (Double.class.equals(field.getType()) || double.class.equals(field.getType())) {
            return true;
        }

        if (Long.class.equals(field.getType()) || long.class.equals(field.getType())) {
            return true;
        }

        if (BigDecimal.class.equals(field.getType())) {
            return true;
        }
        return false;
    }

    public static Boolean isNum(Object cellValue) {
        if (cellValue instanceof Integer
            || cellValue instanceof Double
            || cellValue instanceof Short
            || cellValue instanceof Long
            || cellValue instanceof Float
            || cellValue instanceof BigDecimal) {
            return true;
        }
        return false;
    }

    public static String getDefaultDateString(Date date) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        return simpleDateFormat.format(date);

    }

    public static Date getSimpleDateFormatDate(String value, String format) {
        if (!StringUtils.isEmpty(value)) {
            Date date = null;
            if (!StringUtils.isEmpty(format)) {
                SimpleDateFormat simpleDateFormat = new SimpleDateFormat(format);
                try {
                    date = simpleDateFormat.parse(value);
                    return date;
                } catch (ParseException e) {
                }
            }
            for (String dateFormat : DATE_FORMAT_LIST) {
                try {
                    SimpleDateFormat simpleDateFormat = new SimpleDateFormat(dateFormat);
                    date = simpleDateFormat.parse(value);
                } catch (ParseException e) {
                }
                if (date != null) {
                    break;
                }
            }

            return date;

        }
        return null;

    }


    public static String formatFloat(String value) {
        if (null != value && value.contains(".")) {
            if (isNumeric(value)) {
                try {
                    BigDecimal decimal = new BigDecimal(value);
                    BigDecimal setScale = decimal.setScale(10, BigDecimal.ROUND_HALF_DOWN).stripTrailingZeros();
                    return setScale.toPlainString();
                } catch (Exception e) {
                }
            }
        }
        return value;
    }

    public static String formatFloat0(String value, int n) {
        if (null != value && value.contains(".")) {
            if (isNumeric(value)) {
                try {
                    BigDecimal decimal = new BigDecimal(value);
                    BigDecimal setScale = decimal.setScale(n, BigDecimal.ROUND_HALF_DOWN);
                    return setScale.toPlainString();
                } catch (Exception e) {
                }
            }
        }
        return value;
    }

    public static final Pattern pattern = Pattern.compile("[\\+\\-]?[\\d]+([\\.][\\d]*)?([Ee][+-]?[\\d]+)?$");

    private static boolean isNumeric(String str) {
        Matcher isNum = pattern.matcher(str);
        if (!isNum.matches()) {
            return false;
        }
        return true;
    }

    public static String formatDate(Date cellValue, String format) {
        SimpleDateFormat simpleDateFormat;
        if(!StringUtils.isEmpty(format)) {
             simpleDateFormat = new SimpleDateFormat(format);
        }else {
            simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        }
        return simpleDateFormat.format(cellValue);
    }

    public static String getFieldStringValue(BeanMap beanMap,
                                             String fieldName,
                                             String format,
                                             TypeConverter converter,
                                             List<String> head) {

        Map<String, RowErrorModel.CellInfo> errorMap = (Map<String, RowErrorModel.CellInfo>) beanMap.get("errorMap");

        // 如果是 operateResult 操作结果列
        if (fieldName.equalsIgnoreCase("operateResult")) {
            Integer lineNumber = (Integer) beanMap.get("lineNumber");
            // 如果存在转换器，直接走转换器
            if (converter != null) {
                List<Object> list = new LinkedList<Object>();
                list.add(lineNumber);
                list.add(errorMap);
                return converter.convertToWrite(list);
            }
            return errorMap.toString();
        }

        if (errorMap != null) {
            RowErrorModel.CellInfo cellInfo = errorMap.get(head.get(head.size() - 1));
            if (cellInfo != null) {
                return (String)cellInfo.getSource();
            }
        }

        String cellValue = null;
        Object value = beanMap.get(fieldName);
        if (value != null) {

            // 如果存在转换器，直接走转换器
            if (converter != null) {
                return converter.convertToWrite(value);
            }

            if (value instanceof Date) {
                cellValue = TypeUtil.formatDate((Date)value, format);
            } else {
                cellValue = value.toString();
            }
        }
        return cellValue;
    }

    public static Map getFieldValues(List<String> stringList, ExcelHeadProperty excelHeadProperty, Boolean use1904WindowDate, Map<String, RowErrorModel.CellInfo> rowErrorMap) {
        if (rowErrorMap == null) {
            throw new ExcelAnalysisException(" rowErroMap is null, abort");
        }

        Map map = new HashMap();

        for (int i = 0; i < stringList.size(); i++) {
            ExcelColumnProperty columnProperty = excelHeadProperty.getExcelColumnProperty(i);
            if (columnProperty != null) {
                Object value = null;
                String currentValue = stringList.get(i);
                try {
                    value = TypeUtil.convert(currentValue, columnProperty.getField(),
                            columnProperty.getFormat(), columnProperty.getConverter(), use1904WindowDate);
                } catch (Exception e) {
                    List<String> currentHead = columnProperty.getHead();
                    // TODO 是否需要设置必填项  不需要，解析工具只负责将excel值转成对应bean, 有些业务场景没有必要耦合进来
                    rowErrorMap.put(currentHead.get(currentHead.size() - 1), RowErrorModel.CellInfo.newInstance(currentValue, "格式不正确"));
                }
                if (value != null) {
                    map.put(columnProperty.getField().getName(),value);
                }
            }
        }
        return map;
    }
}
