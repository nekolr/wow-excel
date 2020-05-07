package com.nekolr.annotation;

import com.nekolr.enums.ExcelType;

import java.lang.annotation.*;

/**
 * Excel 文档注解
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface Excel {

    /**
     * Excel 文件名称
     *
     * @return Excel 文件名称
     */
    String value() default "";

    /**
     * Excel 类型
     *
     * @return Excel 文件类型
     * @see ExcelType
     */
    ExcelType type() default ExcelType.XLS;
}
