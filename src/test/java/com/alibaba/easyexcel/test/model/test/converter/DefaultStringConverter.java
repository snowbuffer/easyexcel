package com.alibaba.easyexcel.test.model.test.converter;

import com.alibaba.excel.metadata.TypeConverter;

/**
 * Description:
 *
 * @author cjb
 * @version V1.0
 * @since 2019-03-31 11:32
 */
public class DefaultStringConverter implements TypeConverter {

    @Override
    public Object convertToRead(String input) {
        System.out.println("原值：" + input);
        return "DEFAULT_" + input;
    }

    @Override
    public String convertToWrite(Object value) {
        return null;
    }
}

