package com.nekolr.write.processor;

import com.nekolr.write.ExcelWriteContext;

import java.util.List;

/**
 * 写处理器
 */
public interface ExcelWriteProcessor {

    /**
     * 初始化
     *
     * @param writeContext 写上下文
     */
    void init(ExcelWriteContext writeContext);

    /**
     * 写数据
     *
     * @param data 数据集合
     */
    void write(List<?> data);
}
