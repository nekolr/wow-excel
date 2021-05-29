package com.github.nekolr.write.listener.style;

import com.github.nekolr.metadata.ExcelFieldBean;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

/**
 * 默认的样式
 */
public class DefaultExcelStyleWriteListener implements ExcelStyleWriteListener {

    /**
     * workbook
     */
    private Workbook workbook;

    /**
     * 表格体单元格样式缓存
     */
    private Map<Integer, CellStyle> bodyStyleCache;

    /**
     * 表头单元格样式缓存
     */
    private Map<Integer, CellStyle> headStyleCache;

    @Override
    public void init(Workbook workbook) {
        this.workbook = workbook;
        this.bodyStyleCache = new HashMap<>();
        this.headStyleCache = new HashMap<>();
    }

    @Override
    public void bigTitleStyle(Cell cell) {
        CellStyle cellStyle = this.workbook.createCellStyle();
        // 填充灰色
        cellStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 设置字体样式
        Font font = this.workbook.createFont();
        font.setFontName("宋体");
        font.setBold(true);
        cellStyle.setFont(font);
        // 自动换行
        cellStyle.setWrapText(true);

        cell.setCellStyle(cellStyle);
    }

    @Override
    public void headStyle(Row row, Cell cell, ExcelFieldBean field, Object cellValue, int rowIndex, int colNum) {
        CellStyle cellStyle = this.headStyleCache.get(colNum);
        if (cellStyle == null) {
            cellStyle = this.workbook.createCellStyle();
            // 水平居中
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            // 垂直居中
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            // 自动换行
            cellStyle.setWrapText(true);
            // 填充灰色
            cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 设置边框
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.index);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.index);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setRightBorderColor(IndexedColors.GREY_40_PERCENT.index);
            // 设置字体
            Font font = workbook.createFont();
            font.setBold(true);
            font.setFontName("宋体");
            cellStyle.setFont(font);

            this.headStyleCache.put(colNum, cellStyle);
        }
        cell.setCellStyle(cellStyle);
    }

    @Override
    public void bodyStyle(Row row, Cell cell, ExcelFieldBean field, Object cellValue, int rowIndex, int colNum) {
        // 相同列使用相同的样式，使用缓存可以方便重用同一列的样式
        CellStyle cellStyle = this.bodyStyleCache.get(colNum);
        if (cellStyle == null) {
            cellStyle = this.workbook.createCellStyle();
            // 水平居中
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            // 垂直居中
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            // 自动换行
            cellStyle.setWrapText(true);
            // 设置格式，使用内建格式
            cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(field.getFormat()));

            this.bodyStyleCache.put(colNum, cellStyle);
        }
        cell.setCellStyle(cellStyle);
    }

    @Override
    public void afterWriteCell(Sheet sheet, Row row, Cell cell, ExcelFieldBean field, int rowIndex, int colNum, Object cellValue, boolean isHead) {
        if (isHead) {
            // 首行表头
            if (rowIndex == 0) {
                CellStyle cellStyle = this.workbook.createCellStyle();
                // 设置格式
                cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(field.getFormat()));
                // 表头每列统一使用一个样式
                sheet.setDefaultColumnStyle(colNum, cellStyle);
                sheet.setColumnWidth(colNum, field.getWidth());
            }
            // 设置表头样式
            this.headStyle(row, cell, field, cellValue, rowIndex, colNum);
        } else {
            // 设置表格体样式
            this.bodyStyle(row, cell, field, cellValue, rowIndex, colNum);
        }

    }

    @Override
    public void afterWriteRow(Sheet sheet, Row row, Object rowEntity, int rowIndex, boolean isHead) {

    }
}
