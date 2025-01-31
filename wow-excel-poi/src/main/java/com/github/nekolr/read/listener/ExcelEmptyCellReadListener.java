package com.github.nekolr.read.listener;

import com.github.nekolr.metadata.ExcelFieldBean;

/**
 * 当单元格不存在或者单元格的值为空时，可以使用该监听器处理
 */
@FunctionalInterface
public interface ExcelEmptyCellReadListener extends ExcelReadListener {
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
