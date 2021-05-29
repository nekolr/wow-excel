package com.github.nekolr.convert;

import com.github.nekolr.metadata.DataConverter;
import com.github.nekolr.metadata.ExcelFieldBean;

/**
 * Long 类型和 String 类型的转换器
 * <p>
 * 有时 excel 单元格的格式为文本，但是内容是数值，此时如果实体类的属性使用数值类型将会出错，可以使用此转换器处理
 */
public class StringLongDataConverter implements DataConverter {
    @Override
    public Object toEntityAttribute(ExcelFieldBean field, Object cellValue) {
        return Long.parseLong(cellValue.toString());
    }

    @Override
    public Object toExcelAttribute(ExcelFieldBean field, Object attrValue) {
        return attrValue == null ? null : attrValue.toString();
    }
}
