package com.alibaba.excel.annotation;

import com.alibaba.excel.metadata.CustomCellStyle;
import com.alibaba.excel.metadata.TypeConverter;
import com.alibaba.excel.util.StyleUtil;
import org.apache.poi.ss.usermodel.CellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author jipengfei
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelProperty {

     /**
      * @return
      */
     String[] value() default {""};


     /**
      * @return
      */
     int index() default 99999;

     /**
      *
      * default @see com.alibaba.excel.util.TypeUtil
      * if default is not  meet you can set format
      *
      * @return
      */
     String format() default "";

     /**
      * 自定义类型转换器
      *
      * @return
      */
     Class<? extends TypeConverter> convertor() default TypeConverter.class;

     // 暂时弃用 业务异常场景下，需要业务方指定字段 ，业务方改动大
     Class<? extends CustomCellStyle> errorCellStyle() default CustomCellStyle.class;

     // 暂时弃用 业务异常场景下，需要业务方指定字段 ，业务方改动大
     Class<? extends CustomCellStyle> commonCellStyle() default CustomCellStyle.class;
}
