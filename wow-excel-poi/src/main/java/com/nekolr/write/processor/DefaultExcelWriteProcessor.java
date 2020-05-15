package com.nekolr.write.processor;

import com.nekolr.exception.ExcelReadInitException;
import com.nekolr.exception.ExcelWriteException;
import com.nekolr.exception.ExcelWriteInitException;
import com.nekolr.metadata.*;
import com.nekolr.util.BeanUtils;
import com.nekolr.util.ExcelUtils;
import com.nekolr.util.ParamUtils;
import com.nekolr.write.ExcelWriteContext;
import com.nekolr.write.listener.ExcelWriteEventProcessor;
import com.nekolr.write.listener.style.DefaultExcelStyleWriteListener;
import com.nekolr.write.listener.style.ExcelStyleWriteListener;
import com.nekolr.write.metadata.BigTitle;
import com.nekolr.write.merge.LastCell;
import com.nekolr.write.merge.OldRowCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
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
        this.initStyleWriteListeners(this.writeContext);
    }

    @Override
    public void write(List<?> data) {
        if (data != null && data.size() != 0) {
            Excel excel = this.writeContext.getExcel();
            Sheet sheet = this.getOrCreateSheet();
            int rowNum = this.getRowNum(sheet);
            int colNum = this.writeContext.getColNum();
            List<ExcelField> excelFieldList = excel.getFieldList();
            for (int r = 0; r < data.size(); r++) {
                Object rowEntity = data.get(r);
                Row row = sheet.createRow(rowNum + r);
                for (int col = 0; col < excelFieldList.size(); col++) {
                    ExcelField excelField = excelFieldList.get(col);
                    Cell cell = row.createCell(colNum + col);
                    DataConverter dataConverter = this.getDataConverter(excelField);
                    // 获取属性值
                    Object attrValue = BeanUtils.getFieldValue(rowEntity, excelField.getField());
                    // 使用数据转换器
                    Object cellValue = ExcelUtils.useWriteConverter(attrValue, excelField, dataConverter);
                    // 单元格赋值
                    ExcelUtils.setCellValue(cell, cellValue, excelField);
                    // Event: 写单元格结束后触发
                    ExcelWriteEventProcessor.afterWriteCell(this.writeContext.getCellWriteListeners(), sheet, row, cell, excelField, r, colNum + col, cellValue, false);
                }
                // Event: 写行结束后触发
                ExcelWriteEventProcessor.afterWriteRow(this.writeContext.getRowWriteListeners(), sheet, row, rowEntity, r, false);
            }
        }
    }

    /**
     * 写表头
     */
    @Override
    public void writeHead() {
        LastCell lastCell = null;
        Sheet sheet = this.getOrCreateSheet();
        int rowNum = this.getRowNum(sheet);
        int colNum = this.writeContext.getColNum();
        Excel excel = this.writeContext.getExcel();
        List<ExcelField> excelFieldList = excel.getFieldList();
        // 第一个字段元数据的标题数组的长度就代表了所有表头占用的行数
        for (int r = 0, rowSize = excelFieldList.get(0).getTitles().length; r < rowSize; r++) {
            Row row = sheet.createRow(rowNum + r);
            for (int col = 0, colSize = excelFieldList.size(); col < colSize; col++) {
                ExcelField excelField = excelFieldList.get(col);
                String[] titles = excelField.getTitles();
                String title = titles[r];
                Cell cell = row.createCell(colNum + col);
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
                        ExcelUtils.mergeCols(lastCell, sheet, row, colNum, colNum + col, colSize, title);
                        // 合并行
                        ExcelUtils.mergeRows(oldRowCache, sheet, row, r, colNum + col, rowSize, title);
                    } catch (Exception e) {
                        throw new ExcelWriteException("Auto merge failure", e);
                    }
                }
                // Event: 写单元格结束后触发
                ExcelWriteEventProcessor.afterWriteCell(this.writeContext.getCellWriteListeners(), sheet, row, cell, excelField, r, colNum + col, title, true);
            }
            // Event: 写行结束后触发
            ExcelWriteEventProcessor.afterWriteRow(this.writeContext.getRowWriteListeners(), sheet, row, excelFieldList, r, true);
        }
    }

    @Override
    public void writeBigTitle(BigTitle bigTitle) {
        Sheet sheet = this.getOrCreateSheet();
        int rowNum = this.getRowNum(sheet);
        int endRowNum = rowNum + bigTitle.getLines() - 1;
        for (int r = 0; r < bigTitle.getLines(); r++) {
            Row row = sheet.createRow(rowNum + r);
            // 大标题设置的起始列号和结束列号不受全局设置（写上下文）的 colIndex 影响
            for (int col = bigTitle.getFirstColNum(); col < bigTitle.getLastColNum(); col++) {
                Cell cell = row.createCell(col);
                cell.setCellValue(bigTitle.getContent());
                // Event: 设置大标题的样式
                ExcelWriteEventProcessor.setBigTitleStyle(this.writeContext.getCellWriteListeners(), cell);
            }
        }
        sheet.addMergedRegion(new CellRangeAddress(rowNum, endRowNum, bigTitle.getFirstColNum(), bigTitle.getLastColNum()));
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
                        this.writeContext.setWorkbook(new SXSSFWorkbook(excel.getWindowSize()));
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
            // Event: 在创建 sheet 后触发
            ExcelWriteEventProcessor.afterCreateSheet(this.writeContext.getSheetWriteListeners(), sheet, this.writeContext);
        }
        return sheet;
    }

    /**
     * 计算当前可以开始的行号
     *
     * @param sheet sheet
     * @return 当前可以开始的行号
     */
    private int getRowNum(Sheet sheet) {
        // -1 表示一行都没有
        int lastRowNum = sheet.getLastRowNum();
        int customRowNum = this.writeContext.getRowNum();
        if (lastRowNum == -1) {
            return customRowNum;
        } else {
            return lastRowNum + 1;
        }
    }

    /**
     * 初始化样式监听器
     *
     * @param writeContext 写上下文
     */
    private void initStyleWriteListeners(ExcelWriteContext writeContext) {
        if (writeContext.isUseDefaultStyle()) {
            ExcelStyleWriteListener writeListener = new DefaultExcelStyleWriteListener();
            writeContext.addListener(writeListener);
        }
        writeContext.getCellWriteListeners().forEach(listener -> {
            if (listener instanceof ExcelStyleWriteListener) {
                ((ExcelStyleWriteListener) listener).init(writeContext.getWorkbook());
            }
        });
        writeContext.getRowWriteListeners().forEach(listener -> {
            if (listener instanceof ExcelStyleWriteListener) {
                ((ExcelStyleWriteListener) listener).init(writeContext.getWorkbook());
            }
        });
    }
}
