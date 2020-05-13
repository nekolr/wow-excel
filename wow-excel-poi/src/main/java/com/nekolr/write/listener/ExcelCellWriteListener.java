package com.nekolr.write.listener;

import com.nekolr.metadata.ExcelField;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * 单元格级别的写监听器
 */
@FunctionalInterface
public interface ExcelCellWriteListener extends ExcelWriteListener {

    void afterWriteCell(Sheet sheet, Row row, Cell cell, ExcelField field);
}
