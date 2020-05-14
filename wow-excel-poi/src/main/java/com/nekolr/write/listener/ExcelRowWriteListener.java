package com.nekolr.write.listener;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * 行级别的写监听器
 */
@FunctionalInterface
public interface ExcelRowWriteListener extends ExcelWriteListener {

    /**
     * 在写完行时触发
     *
     * @param sheet     sheet
     * @param row       当前行
     * @param rowEntity 当前行对应的实体
     * @param rowIndex  遍历行时的索引
     */
    void afterWriteRow(Sheet sheet, Row row, Object rowEntity, int rowIndex);
}
