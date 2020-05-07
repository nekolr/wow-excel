package com.nekolr.read.listener;

import com.nekolr.metadata.ExcelFieldBean;
import com.nekolr.metadata.ExcelListener;
import com.nekolr.read.ExcelReadContext;

/**
 * Excel 读监听器
 */
@FunctionalInterface
public interface ExcelReadListener<R> extends ExcelListener {

    /**
     * 在读取一行后调用
     *
     * @param r      一行的数据，对应一个实体对象
     * @param rowNum 行号
     * @return 是否停止读
     */
    boolean afterReadRow(R r, int rowNum);

    /**
     * 在读取一个单元格后调用
     *
     * @param field     字段元数据
     * @param cellValue 单元格值
     * @param rowNum    行号
     * @param colNum    列号
     * @return 处理后的单元格值
     */
    default Object afterReadCell(ExcelFieldBean field, Object cellValue, int rowNum, int colNum) {
        return cellValue;
    }

    /**
     * 在读取完 sheet 后调用
     *
     * @param readContext 读上下文
     */
    default void afterReadSheet(ExcelReadContext<R> readContext) {
    }

    /**
     * 在读 sheet 之前调用
     *
     * @param readContext 读上下文
     */
    default void beforeReadSheet(ExcelReadContext<R> readContext) {
    }

}
