package com.github.nekolr.write.listener;

import com.github.nekolr.write.ExcelWriteContext;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * sheet 级别的写监听器
 */
public interface ExcelSheetWriteListener extends ExcelWriteListener {

    /**
     * 在创建 sheet 之后触发
     *
     * @param sheet        sheet
     * @param writeContext 写上下文
     */
    void afterCreateSheet(Sheet sheet, ExcelWriteContext writeContext);
}
