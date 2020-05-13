package com.nekolr.metadata;

import lombok.Data;

import java.lang.reflect.Field;

/**
 * Excel 字段实体
 */
@Data
public class ExcelField {

    /**
     * 表头标题数组
     */
    private String[] titles;

    /**
     * 表头字段在 Excel 中出现的顺序
     */
    private int order;

    /**
     * 表头字段所在列的数据是否允许出现空值
     */
    private boolean allowEmpty;

    /**
     * 表头字段所在列的数据是否需要自动去除头尾空格
     */
    private boolean autoTrim;

    /**
     * 表头字段所在列的宽度
     */
    private int width;

    /**
     * 表头字段所在列的格式
     */
    private String format;

    /**
     * 是否忽略该字段
     */
    private boolean ignore;

    /**
     * 数据转换器
     */
    private Class<? extends DataConverter> converter;

    /**
     * 使用该注解的 Field
     *
     * @see java.lang.reflect.Field
     */
    private Field field;
}
