package com.nekolr.read.listener;

import com.nekolr.read.ExcelReadContext;

/**
 * sheet 级别的读监听器
 *
 * @param <R> 使用 @Excel 注解的类类型
 */
public interface ExcelSheetReadListener<R> extends ExcelReadListener {
    /**
     * 在读 sheet 之前调用
     *
     * @param readContext 读上下文
     */
    void beforeReadSheet(ExcelReadContext<R> readContext);

    /**
     * 在读完 sheet 后调用
     *
     * @param readContext 读上下文
     */
    void afterReadSheet(ExcelReadContext<R> readContext);
}
