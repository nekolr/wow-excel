package com.github.nekolr.write.listener;

import com.github.nekolr.write.ExcelWriteContext;

/**
 * workbook 级别的写监听器
 */
public interface ExcelWorkbookWriteListener extends ExcelWriteListener {

    /**
     * 在数据刷新到文件或流之前调用
     *
     * @param writeContext 写上下文
     */
    void beforeFlush(ExcelWriteContext writeContext);
}
