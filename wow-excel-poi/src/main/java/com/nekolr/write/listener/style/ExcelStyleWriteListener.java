package com.nekolr.write.listener.style;

import com.nekolr.metadata.ExcelField;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 写样式监听器
 */
public interface ExcelStyleWriteListener {

    void init(Workbook workbook);

    void bigTitleStyle(Cell cell);

    void headStyle(Row row, Cell cell, ExcelField field, String title);

    void bodyStyle(Row row, Cell cell, ExcelField field, String title);
}
