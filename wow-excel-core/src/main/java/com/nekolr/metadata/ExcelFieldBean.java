package com.nekolr.metadata;

import com.nekolr.convert.DataConverter;
import lombok.Data;

import java.lang.reflect.Field;

/**
 * Excel 字段实体
 */
@Data
public class ExcelFieldBean {

    /**
     * 字段名称
     */
    private String filedName;

    /**
     * 表头等级
     * <p>
     * 多级表头是指表头含有上下级关系，如果只有一级上下级关系，则 level = 2
     */
    private int level;

    /**
     * 表头字段在 Excel 中出现的顺序
     */
    private int order;

    /**
     * 表头字段是否允许出现空值
     */
    private boolean allowEmpty;

    /**
     * 字表头段是否需要自动去除头尾空格
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
     */
    private Field excelField;
}
