package com.github.nekolr.read.listener;

/**
 * 行级别的读监听器
 *
 * @param <R> 使用 @Excel 注解的类类型
 */
@FunctionalInterface
public interface ExcelRowReadListener<R> extends ExcelReadListener {

    /**
     * 在读取一行后调用
     *
     * @param r      一行的数据，对应一个实体对象
     * @param rowNum 行号
     * @return 是否停止读
     */
    boolean afterReadRow(R r, int rowNum);

}
