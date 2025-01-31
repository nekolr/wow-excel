package com.github.nekolr.write.listener;

import com.github.nekolr.metadata.ExcelFieldBean;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * 单元格级别的写监听器
 */
@FunctionalInterface
public interface ExcelCellWriteListener extends ExcelWriteListener {

    /**
     * 在写单元格之后触发
     *
     * @param sheet     sheet
     * @param row       当前行
     * @param cell      当前单元格
     * @param field     表头字段元数据
     * @param rowIndex  遍历行时的索引
     * @param colNum    列号
     * @param cellValue 单元格的数据
     * @param isHead    是否是表头单元格
     */
    void afterWriteCell(Sheet sheet, Row row, Cell cell, ExcelFieldBean field, int rowIndex, int colNum, Object cellValue, boolean isHead);
}
