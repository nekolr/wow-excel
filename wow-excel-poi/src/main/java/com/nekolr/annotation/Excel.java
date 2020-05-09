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

    /**
     * 默认情况下，xlsx 文档的 /xl/sharedStrings.xml 数据存储在内存中，这可能会导致内存问题，
     * 使用此选项将此数据存储在临时文件中（H2 database）
     *
     * @return 是否将数据临时存储
     */
    boolean useSstTempFile() default true;

    /**
     * 是否加密临时文件（H2 database）中的数据，如果担心原始数据以明文的形式保存在临时文件，可以启用此选项
     *
     * @return 是否加密此数据
     */
    boolean encryptSstTempFile() default false;
}
