package com.nekolr.metadata;

import com.nekolr.enums.ExcelType;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Excel 实体
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
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
}