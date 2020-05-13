package com.nekolr.write.listener;

import com.nekolr.write.ExcelWriteContext;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * sheet 级别的写监听器
 */
public interface ExcelSheetWriteListener extends ExcelWriteListener {

    void beforeWriteSheet();

    void afterWriteSheet(Sheet sheet, ExcelWriteContext writeContext);
}
