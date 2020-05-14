package com.nekolr.write.listener;

import com.nekolr.metadata.ExcelField;
import com.nekolr.write.ExcelWriteContext;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;

/**
 * Excel 写事件处理器
 */
public class ExcelWriteEventProcessor {

    /**
     * 在写单元格之后触发
     *
     * @param writeListeners 写监听器
     * @param sheet          sheet
     * @param row            当前行
     * @param cell           当前单元格
     * @param field          表头字段元数据
     * @param rowIndex       遍历行时的索引
     * @param colNum         当前列号
     * @param cellValue      单元格的数据
     */
    public static void afterWriteCell(List<ExcelCellWriteListener> writeListeners,
                                      Sheet sheet, Row row, Cell cell, ExcelField field, int rowIndex, int colNum, Object cellValue) {
        writeListeners.forEach(listener -> listener.afterWriteCell(sheet, row, cell, field, rowIndex, colNum, cellValue));
    }

    /**
     * 在写完行之后触发
     *
     * @param writeListeners 写监听器
     * @param sheet          sheet
     * @param row            当前行
     * @param rowEntity      当前行对用的实体数据
     * @param rowIndex       遍历行时的索引
     */
    public static void afterWriteRow(List<ExcelRowWriteListener> writeListeners, Sheet sheet, Row row, Object rowEntity, int rowIndex) {
        writeListeners.forEach(listener -> listener.afterWriteRow(sheet, row, rowEntity, rowIndex));
    }

    /**
     * 在创建 sheet 之后调用
     *
     * @param writeListeners 写监听器集合
     * @param sheet          sheet
     * @param writeContext   写上下文
     */
    public static void afterCreateSheet(List<ExcelSheetWriteListener> writeListeners, Sheet sheet, ExcelWriteContext writeContext) {
        writeListeners.forEach(listener -> listener.afterCreateSheet(sheet, writeContext));
    }

    /**
     * 在数据刷新到文件或流之前调用
     *
     * @param writeListeners 写监听器集合
     * @param writeContext   写上下文
     */
    public static void beforeFlush(List<ExcelWorkbookWriteListener> writeListeners, ExcelWriteContext writeContext) {
        writeListeners.forEach(listener -> listener.beforeFlush(writeContext));
    }
}
