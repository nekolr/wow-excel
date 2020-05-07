package com.nekolr.convert;

import com.nekolr.metadata.DataConverter;
import com.nekolr.metadata.ExcelFieldBean;

/**
 * 默认的转换器，什么也不做
 */
public class DefaultDataConverter implements DataConverter {

    @Override
    public Object toEntityAttribute(ExcelFieldBean field, Object cellValue) {
        return cellValue;
    }

    @Override
    public Object toExcelAttribute(ExcelFieldBean field, Object attrValue) {
        return attrValue;
    }
}