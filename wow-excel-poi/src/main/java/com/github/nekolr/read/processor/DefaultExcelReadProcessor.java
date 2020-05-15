package com.github.nekolr.read.processor;

import com.github.nekolr.Constants;
import com.github.nekolr.exception.ExcelReadException;
import com.github.nekolr.exception.ExcelReadInitException;
import com.github.nekolr.metadata.DataConverter;
import com.github.nekolr.metadata.Excel;
import com.github.nekolr.metadata.ExcelField;
import com.github.nekolr.read.ExcelReadContext;
import com.github.nekolr.read.listener.ExcelEmptyCellReadListener;
import com.github.nekolr.read.listener.ExcelReadEventProcessor;
import com.github.nekolr.util.BeanUtils;
import com.github.nekolr.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 使用 poi 原生的用户模型处理读
 *
 * @param <R> 使用 @Excel 注解的类类型
 */
public class DefaultExcelReadProcessor<R> implements ExcelReadProcessor<R> {

    /**
     * 读上下文
     */
    private ExcelReadContext<R> readContext;

    @Override
    public void init(ExcelReadContext<R> readContext) {
        this.readContext = readContext;
        this.createWorkbook(this.readContext);
    }

    @Override
    public void read() {
        List<R> resultList;
        if (this.readContext.isAllSheets()) {
            resultList = this.readAllSheets();
        } else {
            resultList = this.readSheet();
        }
        // Event: 所有要求的任务都执行完毕后触发，结果通知
        ExcelReadEventProcessor.resultNotify(this.readContext.getReadResultListener(), resultList);
    }

    /**
     * 执行读
     *
     * @param sheet worksheet
     */
    private List<R> doRead(Sheet sheet) {
        // Event: 开始读之前触发
        ExcelReadEventProcessor.beforeReadSheet(this.readContext.getSheetReadListeners(), this.readContext);
        Excel excel = this.readContext.getExcel();
        List<ExcelField> excelFieldList = excel.getFieldList();
        int actualRowNum = this.readContext.getRowNum();
        int colNum = this.readContext.getColNum();
        boolean saveResult = this.readContext.isSaveResult();
        R r;
        Object cellValue;
        boolean isStop = false;
        List<R> resultList;
        if (saveResult) {
            resultList = new ArrayList<>(sheet.getLastRowNum());
        } else {
            resultList = new ArrayList<>(0);
        }
        for (Row row : sheet) {
            if (isStop) {
                break;
            }
            if (row.getRowNum() >= actualRowNum) {
                try {
                    r = this.readContext.getExcelClass().newInstance();
                } catch (InstantiationException | IllegalAccessException e) {
                    throw new ExcelReadInitException("Excel entity init failure: " + e.getMessage());
                }
                for (int col = 0; col < excelFieldList.size(); col++) {
                    ExcelField excelField = excelFieldList.get(col);
                    if (!excelField.isIgnore()) {
                        Field field = excelFieldList.get(col).getField();
                        Cell cell = row.getCell(colNum + col);
                        DataConverter dataConverter = this.getDataConverter(excelField);
                        if (cell != null) {
                            cellValue = ExcelUtils.getCellValue(cell, excelField, field);
                            if (cellValue != null) {
                                // Event: 读单元格结束后触发
                                cellValue = ExcelReadEventProcessor.afterReadCell(this.readContext.getCellReadListeners(), excelField, cellValue, row.getRowNum(), col);
                                // 使用转换器转换结果
                                cellValue = ExcelUtils.useReadConverter(cellValue, excelField, dataConverter);
                            } else {
                                // 单元格值为空
                                if (!excelField.isAllowEmpty()) {
                                    // 如果字段不允许空值，那么需要用户通过监听器来自己处理空值
                                    cellValue = this.handleEmpty(this.readContext, excelField, row.getRowNum(), col);
                                }
                            }
                        } else {
                            // 单元格不存在
                            if (!excelField.isAllowEmpty()) {
                                cellValue = this.handleEmpty(this.readContext, excelField, row.getRowNum(), col);
                            } else {
                                cellValue = null;
                            }
                        }
                        // 设置字段的值
                        this.setFieldValue(r, field, cellValue);
                    }
                }
                // Event: 读行结束后触发
                isStop = ExcelReadEventProcessor.afterReadRow(this.readContext.getRowReadListeners(), r, row.getRowNum());
                // 将结果放入集合
                if (saveResult) resultList.add(r);
            }
        }
        // Event: sheet 读取完毕后触发
        ExcelReadEventProcessor.afterReadSheet(this.readContext.getSheetReadListeners(), this.readContext);
        return resultList;
    }

    /**
     * 单元格为空时执行
     *
     * @param readContext 读上下文
     * @param excelField  字段元数据实体
     * @param rowNum      行号
     * @param colNum      列号
     */
    private Object handleEmpty(ExcelReadContext<R> readContext, ExcelField excelField, int rowNum, int colNum) {
        List<ExcelEmptyCellReadListener> emptyReadListeners = readContext.getEmptyCellReadListeners();
        if (emptyReadListeners == null) {
            throw new ExcelReadInitException("If the field is not allowed empty, please specify a listener that handles null value.");
        }
        return ExcelReadEventProcessor.afterReadEmptyCell(emptyReadListeners, excelField, rowNum, colNum);
    }

    /**
     * 获取数据转换器
     *
     * @param excelField 表头字段对应的元数据
     * @return 数据转换器
     */
    private DataConverter getDataConverter(ExcelField excelField) {
        Class<? extends DataConverter> converterClass = excelField.getConverter();
        DataConverter dataConverter = this.readContext.getConverterCache().get(converterClass);
        if (dataConverter == null) {
            try {
                DataConverter converter = excelField.getConverter().newInstance();
                this.readContext.getConverterCache().put(converterClass, converter);
                return converter;
            } catch (InstantiationException | IllegalAccessException e) {
                throw new ExcelReadInitException("Data converter: " + converterClass.getName() + " init failure: " + e.getMessage());
            }
        }
        return dataConverter;
    }

    /**
     * 创建 workbook
     *
     * @param readContext 读上下文
     * @return Workbook
     */
    private Workbook createWorkbook(ExcelReadContext<R> readContext) {
        Workbook workbook = readContext.getWorkbook();
        if (workbook == null) {
            if (readContext.getFile() != null) {
                try {
                    workbook = WorkbookFactory.create(readContext.getFile(), readContext.getPassword());
                    readContext.setWorkbook(workbook);
                } catch (Exception e) {
                    throw new ExcelReadException("Can not read this workbook: " + e.getMessage(), e);
                }
            } else {
                try {
                    workbook = WorkbookFactory.create(readContext.getInputStream(), readContext.getPassword());
                    readContext.setWorkbook(workbook);
                } catch (Exception e) {
                    throw new ExcelReadException("Can not read this workbook: " + e.getMessage(), e);
                }
            }
        }
        return workbook;
    }

    /**
     * 读所有的 sheets
     *
     * @return 结果集合
     */
    private List<R> readAllSheets() {
        List<R> resultList = new ArrayList<>();
        for (Sheet sheet : this.readContext.getWorkbook()) {
            resultList.addAll(this.doRead(sheet));
        }
        return resultList;
    }

    /**
     * 读指定的 sheet
     *
     * @return 结果结合
     */
    private List<R> readSheet() {
        Sheet sheet;
        if (this.readContext.getSheetName() != null) {
            sheet = this.readContext.getWorkbook().getSheet(this.readContext.getSheetName());
        } else if (this.readContext.getSheetAt() != null) {
            sheet = this.readContext.getWorkbook().getSheetAt(this.readContext.getSheetAt());
        } else {
            sheet = this.readContext.getWorkbook().getSheetAt(Constants.DEFAULT_SHEET_AT);
        }
        return this.doRead(sheet);
    }

    /**
     * 设置字段的值
     *
     * @param r         实体对象
     * @param field     字段
     * @param cellValue 单元格的值
     */
    private void setFieldValue(R r, Field field, Object cellValue) {
        if (field.getType() == LocalDate.class) {
            BeanUtils.setFieldValue(r, field, LocalDateTime.ofInstant(((Date) cellValue).toInstant(), ZoneId.systemDefault()).toLocalDate());
            return;
        }
        if (field.getType() == LocalDateTime.class) {
            BeanUtils.setFieldValue(r, field, LocalDateTime.ofInstant(((Date) cellValue).toInstant(), ZoneId.systemDefault()));
            return;
        }
        try {
            BeanUtils.setFieldValue(r, field, cellValue);
        } catch (RuntimeException e) {
            throw new IllegalArgumentException("Unsupported data type, field: " + field.getName() + " value: " + cellValue);
        }
    }
}
