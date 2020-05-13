package com.nekolr.write.listener;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * 行级别的写监听器
 */
@FunctionalInterface
public interface ExcelRowWriteListener extends ExcelWriteListener {

    void afterWriteRow(Sheet sheet, Row row, Object rowValue);
}
