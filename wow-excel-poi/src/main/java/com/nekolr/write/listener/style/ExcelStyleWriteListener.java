package com.nekolr.write.listener.style;

import com.nekolr.metadata.ExcelField;
import com.nekolr.write.listener.ExcelCellWriteListener;
import com.nekolr.write.listener.ExcelRowWriteListener;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 写样式监听器
 */
public interface ExcelStyleWriteListener extends ExcelRowWriteListener, ExcelCellWriteListener {

    /**
     * 初始化
     *
     * @param workbook workbook
     */
    void init(Workbook workbook);

    /**
     * 设置大标题的样式
     *
     * @param cell 大标题所在的单元格
     */
    default void bigTitleStyle(Cell cell) {
    }

    /**
     * 设置表头的样式
     *
     * @param row       当前行
     * @param cell      当前单元格
     * @param field     表头字段元数据
     * @param cellValue 单元格的值
     * @param rowIndex  遍历行时的索引
     * @param colNum    列号
     */
    default void headStyle(Row row, Cell cell, ExcelField field, Object cellValue, int rowIndex, int colNum) {
    }

    /**
     * 设置表格主体的样式
     *
     * @param row       当前行
     * @param cell      当前单元格
     * @param field     表头字段元数据
     * @param cellValue 单元格的值
     * @param rowIndex  遍历行时的索引
     * @param colNum    列号
     */
    void bodyStyle(Row row, Cell cell, ExcelField field, Object cellValue, int rowIndex, int colNum);
}
