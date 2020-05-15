package com.github.nekolr.metadata;

/**
 * 数据转换接口
 */
public interface DataConverter {

    /**
     * 将单元格的值按照实体属性的类型进行转换
     *
     * @param field     字段元数据
     * @param cellValue 单元格的值
     * @return 转换后的值
     */
    Object toEntityAttribute(ExcelField field, Object cellValue);

    /**
     * 将实体属性的值按照 excel 单元格的数据类型进行转换
     *
     * @param field     字段元数据
     * @param attrValue 实体属性的值
     * @return 转换后的值
     */
    Object toExcelAttribute(ExcelField field, Object attrValue);
}
