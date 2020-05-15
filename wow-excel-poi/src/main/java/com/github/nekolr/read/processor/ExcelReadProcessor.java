package com.github.nekolr.read.processor;

import com.github.nekolr.read.ExcelReadContext;

/**
 * Excel 读处理器
 *
 * @param <R> 使用 @Excel 注解的类类型
 */
public interface ExcelReadProcessor<R> {

    /**
     * 初始化
     *
     * @param readContext 读上下文
     */
    void init(ExcelReadContext<R> readContext);

    /**
     * 读
     */
    void read();
}
