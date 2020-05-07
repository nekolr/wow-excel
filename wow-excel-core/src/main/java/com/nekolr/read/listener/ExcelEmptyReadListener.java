package com.nekolr.read.listener;

import com.nekolr.metadata.ExcelFieldBean;
import com.nekolr.metadata.ExcelListener;

/**
 * 当单元格不存在或者单元格的值为空时
 */
@FunctionalInterface
public interface ExcelEmptyReadListener extends ExcelListener {
    /**
     * 处理空值
     *
     * @param field  字段元数据
     * @param rowNum 行号
     * @param colNum 列号
     * @return 处理后的值
     */
    Object handleEmpty(ExcelFieldBean field, int rowNum, int colNum);
}
