package com.nekolr.write.processor;

import com.nekolr.write.ExcelWriteContext;
import com.nekolr.write.metadata.BigTitle;

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

    /**
     * 写表头
     */
    void writeHead();

    /**
     * 写大标题
     */
    default void writeBigTitle(BigTitle bigTitle) {
    }
}
