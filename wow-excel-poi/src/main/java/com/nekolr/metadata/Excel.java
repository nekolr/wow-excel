package com.nekolr.metadata;

import com.nekolr.enums.ExcelType;
import lombok.Data;

import java.util.List;

/**
 * Excel 实体
 */
@Data
public class Excel {

    /**
     * Excel 名称
     */
    private String excelName;

    /**
     * Excel 类型
     *
     * @see ExcelType
     */
    private ExcelType excelType;

    /**
     * 字段列表
     *
     * @see ExcelField
     */
    private List<ExcelField> fieldList;

    /**
     * 加载到并常驻内存中的数据行数
     */
    private int rowCacheSize;

    /**
     * 将输入流读取到磁盘时所使用的缓冲区大小，单位 byte
     */
    private int bufferSize;

    /**
     * 默认情况下，xlsx 文档的 /xl/sharedStrings.xml 数据存储在内存中，这可能会导致内存问题，
     * 使用此选项将此数据存储在临时文件中（H2 database）
     */
    private boolean useSstTempFile;

    /**
     * 是否加密临时文件（H2 database）中的数据，如果担心原始数据以明文的形式保存在临时文件，可以启用此选项
     */
    private boolean encryptSstTempFile;
}