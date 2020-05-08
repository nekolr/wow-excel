package com.nekolr.convert;

import com.nekolr.metadata.DataConverter;
import com.nekolr.metadata.ExcelField;

/**
 * 默认的转换器，什么也不做
 */
public class DefaultDataConverter implements DataConverter {

    @Override
    public Object toEntityAttribute(ExcelField field, Object cellValue) {
        return cellValue;
    }

    @Override
    public Object toExcelAttribute(ExcelField field, Object attrValue) {
        return attrValue;
    }
}
