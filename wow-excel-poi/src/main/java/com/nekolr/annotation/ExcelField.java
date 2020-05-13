package com.nekolr.annotation;

import com.nekolr.metadata.DataConverter;
import com.nekolr.convert.DefaultDataConverter;

import java.lang.annotation.*;

/**
 * Excel 表头字段
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelField {

    /**
     * 表头字段名称
     *
     * @return 表头字段名称
     */
    String[] value() default {""};

    /**
     * 字段出现在 Excel 中的顺序，数值小的在前，默认情况下为实体字段顺序
     *
     * @return 顺序值
     */
    int order() default 0;

    /**
     * 字段是否允许出现空值
     *
     * @return 字段是否允许出现空值
     */
    boolean allowEmpty() default true;

    /**
     * 字段是否需要自动去除首尾空格
     *
     * @return 字段是否需要自动去除首尾空格
     */
    boolean autoTrim() default true;

    /**
     * 表头所在列的宽度，默认为 8 个字符的宽度，也就是 8 * 256
     * <p>
     * 列宽是通过字符个数来确定的，每列可显示的最大字符数为 255，列宽单位为一个字符宽度的 1/256
     *
     * @return 表头所在列的宽度
     */
    int width() default 2048;

    /**
     * 单元格格式，比如日期格式
     *
     * @return 单元格格式
     */
    String format() default "";

    /**
     * 是否忽略此字段
     *
     * @return 是否忽略此字段
     */
    boolean ignore() default false;

    /**
     * 数据转换器，提供数据转换功能
     *
     * @return 数据转换器
     */
    Class<? extends DataConverter> converter() default DefaultDataConverter.class;

}
