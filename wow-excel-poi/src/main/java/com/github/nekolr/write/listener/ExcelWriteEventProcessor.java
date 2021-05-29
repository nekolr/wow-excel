package com.github.nekolr.write.listener;

import com.github.nekolr.write.ExcelWriteContext;
import com.github.nekolr.metadata.ExcelFieldBean;
import com.github.nekolr.write.listener.style.ExcelStyleWriteListener;
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
     * @param isHead         是否是表头单元格
     */
    public static void afterWriteCell(List<ExcelCellWriteListener> writeListeners,
                                      Sheet sheet, Row row, Cell cell, ExcelFieldBean field, int rowIndex, int colNum, Object cellValue, boolean isHead) {
        writeListeners.forEach(listener -> listener.afterWriteCell(sheet, row, cell, field, rowIndex, colNum, cellValue, isHead));
    }

    /**
     * 在写完行之后触发
     *
     * @param writeListeners 写监听器
     * @param sheet          sheet
     * @param row            当前行
     * @param rowEntity      当前行对用的实体数据
     * @param rowIndex       遍历行时的索引
     * @param isHead         是否是表头所在行
     */
    public static void afterWriteRow(List<ExcelRowWriteListener> writeListeners, Sheet sheet, Row row, Object rowEntity, int rowIndex, boolean isHead) {
        writeListeners.forEach(listener -> listener.afterWriteRow(sheet, row, rowEntity, rowIndex, isHead));
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

    /**
     * 设置大标题的样式
     *
     * @param writeListeners 写监听器集合
     * @param cell           大标题所在单元格
     */
    public static void setBigTitleStyle(List<ExcelCellWriteListener> writeListeners, Cell cell) {
        writeListeners.forEach(listener -> {
            if (listener instanceof ExcelStyleWriteListener) {
                ((ExcelStyleWriteListener) listener).bigTitleStyle(cell);
            }
        });
    }
}
