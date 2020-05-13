package com.nekolr.write.processor;

import com.nekolr.exception.ExcelReadInitException;
import com.nekolr.exception.ExcelWriteException;
import com.nekolr.exception.ExcelWriteInitException;
import com.nekolr.metadata.*;
import com.nekolr.util.BeanUtils;
import com.nekolr.util.ExcelUtils;
import com.nekolr.util.ParamUtils;
import com.nekolr.write.ExcelWriteContext;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 默认的写处理器
 */
public class DefaultExcelWriteProcessor implements ExcelWriteProcessor {

    /**
     * 写上下文
     */
    private ExcelWriteContext writeContext;

    /**
     * 旧的行数据缓存
     */
    private Map<Integer, OldRowCell> oldRowCache;

    @Override
    public void init(ExcelWriteContext writeContext) {
        this.writeContext = writeContext;
        this.createWorkbook();
    }

    @Override
    public void write(List<?> data) {
        if (data != null && data.size() != 0) {
            Excel excel = this.writeContext.getExcel();
            Sheet sheet = this.getOrCreateSheet();
            int rowIndex = this.getRowIndex(sheet);
            int colIndex = this.writeContext.getColIndex();
            List<ExcelField> excelFieldList = excel.getFieldList();
            for (int i = 0; i < data.size(); i++) {
                Object entity = data.get(i);
                Row row = sheet.createRow(rowIndex);
                for (int col = 0; col < excelFieldList.size(); col++) {
                    ExcelField excelField = excelFieldList.get(col);
                    Cell cell = row.createCell(colIndex + col);
                    DataConverter dataConverter = this.getDataConverter(excelField);
                    // 获取属性值
                    Object attrValue = BeanUtils.getFieldValue(entity, excelField.getField());
                    // 使用数据转换器
                    attrValue = ExcelUtils.useWriteConverter(attrValue, excelField, dataConverter);
                    // 单元格赋值
                    ExcelUtils.setCellValue(cell, attrValue, excelField);
                }
            }
        }
    }

    /**
     * 写表头
     * TODO: 没给样式
     */
    @Override
    public void writeHead() {
        LastCell lastCell = null;
        Sheet sheet = this.getOrCreateSheet();
        int rowIndex = this.getRowIndex(sheet);
        int colIndex = this.writeContext.getColIndex();
        Excel excel = this.writeContext.getExcel();
        List<ExcelField> excelFieldList = excel.getFieldList();
        // 第一个字段元数据的标题数组的长度就代表了所有表头占用的行数
        for (int r = 0, rowSize = excelFieldList.get(0).getTitles().length; r < rowSize; r++) {
            Row row = sheet.createRow(rowIndex + r);
            for (int col = 0, colSize = excelFieldList.size(); col < colSize; col++) {
                String[] titles = excelFieldList.get(col).getTitles();
                String title = titles[r];
                Cell cell = row.createCell(colIndex + col);
                // 设置表头名称
                cell.setCellValue(title);
                // 需要写多级表头
                if (this.writeContext.isMultiHead()) {
                    if (lastCell == null) {
                        lastCell = new LastCell();
                    }
                    if (this.oldRowCache == null) {
                        this.oldRowCache = new HashMap<>();
                    }
                    try {
                        // 合并列
                        ExcelUtils.mergeCols(lastCell, sheet, row, colIndex, colIndex + col, colSize, title);
                        // 合并行
                        ExcelUtils.mergeRows(oldRowCache, sheet, row, r, colIndex + col, rowSize, title);
                    } catch (Exception e) {
                        throw new ExcelWriteException("Auto merge failure", e);
                    }
                }
            }
        }

    }

    @Override
    public void writeBigTitle() {
        Sheet sheet = this.getOrCreateSheet();
    }

    /**
     * 获取数据转换器
     *
     * @param excelField 表头字段对应的元数据
     * @return 数据转换器
     */
    private DataConverter getDataConverter(ExcelField excelField) {
        Class<? extends DataConverter> converterClass = excelField.getConverter();
        DataConverter dataConverter = this.writeContext.getConverterCache().get(converterClass);
        if (dataConverter == null) {
            try {
                DataConverter converter = excelField.getConverter().newInstance();
                this.writeContext.getConverterCache().put(converterClass, converter);
                return converter;
            } catch (InstantiationException | IllegalAccessException e) {
                throw new ExcelReadInitException("Data converter: " + converterClass.getName() + " init failure: " + e.getMessage());
            }
        }
        return dataConverter;
    }

    /**
     * 创建 Workbook
     *
     * @return Workbook
     */
    private void createWorkbook() {
        Workbook workbook = this.writeContext.getWorkbook();
        if (workbook == null) {
            Excel excel = this.writeContext.getExcel();
            switch (excel.getWorkbookType()) {
                case XLS:
                    this.writeContext.setWorkbook(new HSSFWorkbook());
                    break;
                case XLSX:
                    if (this.writeContext.isStreamingWriterEnabled()) {
                        this.writeContext.setWorkbook(new SXSSFWorkbook());
                    } else {
                        this.writeContext.setWorkbook(new XSSFWorkbook());
                    }
                    break;
                default:
                    throw new ExcelWriteInitException("Unsupported format: " + excel.getWorkbookType().name());
            }
        }
    }

    /**
     * 创建 Sheet
     *
     * @return Sheet
     */
    private Sheet getOrCreateSheet() {
        Workbook workbook = this.writeContext.getWorkbook();
        String sheetName = this.writeContext.getSheetName();
        Sheet sheet = this.writeContext.getSheet();
        if (sheet == null) {
            if (ParamUtils.isEmpty(sheetName)) {
                sheet = workbook.createSheet();
            } else {
                sheet = workbook.createSheet(sheetName);
            }
            this.writeContext.setSheet(sheet);
        }
        return sheet;
    }

    /**
     * 计算当前可以开始的行
     *
     * @param sheet sheet
     * @return 当前可以开始的行
     */
    private int getRowIndex(Sheet sheet) {
        // -1 表示一行都没有
        int lastRowIndex = sheet.getLastRowNum();
        int customIndex = this.writeContext.getRowIndex();
        if (lastRowIndex == -1) {
            return customIndex;
        } else {
            return lastRowIndex + 1;
        }
    }
}
