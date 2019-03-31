package com.alibaba.excel.metadata;

/**
 * 类型转换器
 */
public interface TypeConverter {

   /**
    * excel值转javaBean
    *
    * @param input
    * @return
    */
   Object convertToRead(String input);

   /**
    * javaBean转excel
    *
    * @param value
    * @return
    */
   String convertToWrite(Object value);

}