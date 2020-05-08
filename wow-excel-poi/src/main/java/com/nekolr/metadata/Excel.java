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
}