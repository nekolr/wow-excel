package com.nekolr.read.listener;

import com.nekolr.metadata.ExcelField;

/**
 * 单元格级别的读监听器
 */
@FunctionalInterface
public interface ExcelCellReadListener extends ExcelReadListener {

    /**
     * 在读取一个单元格后调用
     *
     * @param field     字段元数据
     * @param cellValue 单元格值
     * @param rowNum    行号
     * @param colNum    列号
     * @return 处理后的单元格值
     */
    Object afterReadCell(ExcelField field, Object cellValue, int rowNum, int colNum);
}
