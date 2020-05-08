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

    /**
     * 加载到并常驻内存中的数据行数
     *
     * @return 加载到内存中的行数
     */
    int rowCacheSize() default 100;

    /**
     * 将输入流读取到磁盘时所使用的缓冲区大小，单位 byte
     *
     * @return 缓冲区大小
     */
    int bufferSize() default 2048;
}
