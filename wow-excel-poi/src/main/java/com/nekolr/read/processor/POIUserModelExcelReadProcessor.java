package com.nekolr.read.processor;

import com.nekolr.exception.ExcelReadInitException;
import com.nekolr.exception.ExcelReadException;
import com.nekolr.metadata.DataConverter;
import com.nekolr.metadata.ExcelBean;
import com.nekolr.metadata.ExcelFieldBean;
import com.nekolr.metadata.ExcelListener;
import com.nekolr.read.ExcelReadContext;
import com.nekolr.read.listener.ExcelEmptyReadListener;
import com.nekolr.read.listener.ExcelReadEventProcessor;
import com.nekolr.read.listener.ExcelReadListener;
import com.nekolr.util.BeanUtils;
import com.nekolr.util.ExcelUtils;
import com.nekolr.util.ParameterUtils;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * 使用 poi 原生方式处理读
 *
 * @param <R> 使用 @Excel 注解的类类型
 */
public class POIUserModelExcelReadProcessor<R> implements ExcelReadProcessor<R> {

    /**
     * 读上下文
     */
    private ExcelReadContext<R> readContext;

    @Override
    public void init(ExcelReadContext<R> readContext) {
        this.readContext = readContext;
        Workbook workbook;
        try {
            workbook = WorkbookFactory.create(this.readContext.getInputStream(), this.readContext.getPassword());
            this.readContext.setWorkbook(workbook);
        } catch (Exception e) {
            throw new ExcelReadException(e.getMessage());
        }
    }

    @Override
    public void read(int headIndex, int colIndex, String sheetName) {
        Sheet sheet = this.readContext.getWorkbook().getSheet(sheetName);
        if (sheet == null) {
            throw new ExcelReadException("The " + sheetName + " is not found in the workbook");
        }
        this.doRead(headIndex, colIndex, sheet);
    }

    @Override
    public void read(int headIndex, int colIndex, int sheetAt) {
        Sheet sheet = this.readContext.getWorkbook().getSheetAt(sheetAt);
        if (sheet == null) {
            throw new ExcelReadException("The sheetAt: " + sheetAt + " is not found in the workbook");
        }
        this.doRead(headIndex, colIndex, sheet);
    }

    /**
     * 执行读
     *
     * @param headIndex 表头开始位置
     * @param colIndex  列开始位置
     * @param sheet     worksheet
     */
    private void doRead(int headIndex, int colIndex, Sheet sheet) {
        List<ExcelListener> readListeners = this.readContext.getReadListenerCache().get(ExcelReadListener.class);
        // Event: 开始读之前触发
        ExcelReadEventProcessor.beforeReadSheet(readListeners, this.readContext);
        ExcelBean excelBean = this.readContext.getExcelBean();
        List<ExcelFieldBean> excelFieldBeanList = excelBean.getFieldList();
        // 计算数据行的起始位置
        int rowIndex = ParameterUtils.checkHeadAndGetStartIndex(excelBean);
        // 数据行实际的起始位置
        int actualRowIndex = headIndex + rowIndex;
        R r;
        Object cellValue;
        boolean isStop = false;
        List<R> resultList = new ArrayList<>(sheet.getLastRowNum());
        for (Row row : sheet) {
            if (isStop) {
                break;
            }
            if (row.getRowNum() >= actualRowIndex) {
                try {
                    r = this.readContext.getExcelClass().newInstance();
                } catch (InstantiationException | IllegalAccessException e) {
                    throw new ExcelReadInitException("Excel entity init failure: " + e.getMessage());
                }
                for (int col = 0; col < excelFieldBeanList.size(); col++) {
                    Field field = excelFieldBeanList.get(col).getExcelField();
                    Cell cell = row.getCell(colIndex + col);
                    ExcelFieldBean excelFieldBean = excelFieldBeanList.get(col);
                    DataConverter dataConverter = this.getDataConverter(excelFieldBean);
                    if (cell != null) {
                        cellValue = ExcelUtils.getCellValue(cell, excelFieldBean, field);
                        if (cellValue != null) {
                            // Event: 读单元格结束后触发
                            cellValue = ExcelReadEventProcessor.afterReadCell(readListeners, excelFieldBean, cellValue, row.getRowNum(), col);
                            // 使用转换器转换结果
                            cellValue = ExcelUtils.useReadConverter(cellValue, excelFieldBean, dataConverter);
                        } else {
                            // 单元格值为空
                            if (!excelFieldBean.isAllowEmpty()) {
                                // 如果字段不允许空值，那么需要用户通过监听器来自己处理空值
                                cellValue = this.handleEmpty(this.readContext, excelFieldBean, row.getRowNum(), col);
                            }
                        }
                    } else {
                        // 单元格不存在
                        if (!excelFieldBean.isAllowEmpty()) {
                            cellValue = this.handleEmpty(this.readContext, excelFieldBean, row.getRowNum(), col);
                        } else {
                            cellValue = null;
                        }
                    }
                    // 设置字段的值
                    BeanUtils.setFieldValue(r, field, cellValue);
                }
                // Event: 读行结束后触发
                isStop = ExcelReadEventProcessor.afterReadRow(readListeners, r, row.getRowNum());
                // 将结果放入集合
                resultList.add(r);
            }
        }
        // Event: sheet 读取完毕后触发
        ExcelReadEventProcessor.afterReadSheet(readListeners, this.readContext);
        // Event: sheet 读取完毕后触发，结果通知
        ExcelReadEventProcessor.resultNotify(this.readContext.getReadResultListener(), resultList);
    }

    /**
     * 单元格为空时执行
     *
     * @param readContext    读上下文
     * @param excelFieldBean 字段元数据实体
     * @param rowNum         行号
     * @param colNum         列号
     */
    private Object handleEmpty(ExcelReadContext<R> readContext, ExcelFieldBean excelFieldBean, int rowNum, int colNum) {
        List<ExcelListener> emptyReadListeners = readContext.getReadListenerCache().get(ExcelEmptyReadListener.class);
        if (emptyReadListeners == null) {
            throw new ExcelReadInitException("If the field is not allowed empty, please specify a listener that handles null value.");
        }
        return ExcelReadEventProcessor.afterReadEmptyCell(emptyReadListeners, excelFieldBean, rowNum, colNum);
    }

    /**
     * 获取数据转换器
     *
     * @param excelFieldBean 表头字段对应的元数据
     * @return 数据转换器
     */
    private DataConverter getDataConverter(ExcelFieldBean excelFieldBean) {
        Class<? extends DataConverter> converterClass = excelFieldBean.getConverter();
        DataConverter dataConverter = this.readContext.getConverterCache().get(converterClass);
        if (dataConverter == null) {
            try {
                DataConverter converter = excelFieldBean.getConverter().newInstance();
                this.readContext.getConverterCache().put(converterClass, converter);
                return converter;
            } catch (InstantiationException | IllegalAccessException e) {
                throw new ExcelReadInitException("Data converter: " + converterClass.getName() + " init failure: " + e.getMessage());
            }
        }
        return dataConverter;
    }
}
