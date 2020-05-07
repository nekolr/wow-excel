package com.nekolr.read.processor;

import com.nekolr.read.ExcelReadContext;

/**
 * Excel 读处理器
 *
 * @param <R> 使用 @Excel 注解的类类型
 */
public interface ExcelReadProcessor<R> {

    /**
     * 初始化
     *
     * @param excelReadContext 读上下文
     */
    void init(ExcelReadContext<R> excelReadContext);

    /**
     * 读
     *
     * @param headIndex 开始行
     * @param colIndex  开始列
     * @param sheetName sheet 名称
     */
    void read(int headIndex, int colIndex, String sheetName);

    /**
     * 读
     *
     * @param headIndex 开始行
     * @param colIndex  开始列
     * @param sheetAt   sheet 位置
     */
    void read(int headIndex, int colIndex, int sheetAt);
}
