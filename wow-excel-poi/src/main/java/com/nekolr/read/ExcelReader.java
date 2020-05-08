package com.nekolr.read;

import com.nekolr.metadata.ExcelListener;
import com.nekolr.metadata.ExcelReadResultListener;
import com.nekolr.read.listener.ExcelReadListener;
import com.nekolr.read.processor.DefaultExcelReadProcessor;
import com.nekolr.read.processor.ExcelReadProcessor;

import java.io.IOException;
import java.util.*;

/**
 * ExcelReader
 *
 * @param <R> 使用 @Excel 注解的类类型
 */
public class ExcelReader<R> {

    /**
     * 读上下文
     */
    private ExcelReadContext<R> readContext;

    /**
     * 读处理器
     */
    private ExcelReadProcessor<R> excelReadProcessor;


    public ExcelReader(ExcelReadContext<R> readContext) {
        this.readContext = readContext;
        this.excelReadProcessor = this.lookupReadProcessor();
    }

    /**
     * 寻找读处理器
     *
     * @return ExcelReadProcessor
     */
    private ExcelReadProcessor<R> lookupReadProcessor() {
        ServiceLoader<ExcelReadProcessor> processors = ServiceLoader.load(ExcelReadProcessor.class);
        for (ExcelReadProcessor<R> processor : processors) {
            return processor;
        }
        // 在找不到其他处理器的情况下返回默认的处理器
        return new DefaultExcelReadProcessor<>();
    }

    /**
     * 设置文档密码
     *
     * @param password 密码
     * @return ExcelReader
     */
    public ExcelReader<R> password(String password) {
        this.readContext.setPassword(password);
        return this;
    }

    /**
     * 设置 sheet 名称
     *
     * @param sheetName sheet 名称
     * @return ExcelReader
     */
    public ExcelReader<R> sheetName(String sheetName) {
        this.readContext.setSheetName(sheetName);
        return this;
    }

    /**
     * 设置 sheet 坐标
     *
     * @param sheetAt sheet 坐标
     * @return ExcelReader
     */
    public ExcelReader<R> sheetAt(Integer sheetAt) {
        if (sheetAt < 0) {
            throw new IllegalArgumentException("Sheet index must be greater than or equal to 0");
        }
        this.readContext.setSheetAt(sheetAt);
        return this;
    }

    /**
     * 设置数据起始行号
     * <p>
     * 是数据开始的行号，不是表头开始的行号
     *
     * @param rowIndex 数据起始行号
     * @return ExcelReader
     */
    public ExcelReader<R> rowIndex(int rowIndex) {
        if (rowIndex < 0) {
            throw new IllegalArgumentException("Row index must be greater than or equal to 0");
        }
        this.readContext.setRowIndex(rowIndex);
        return this;
    }

    /**
     * 设置数据起始列号
     *
     * @param colIndex 数据起始行号
     * @return ExcelReader
     */
    public ExcelReader<R> colIndex(int colIndex) {
        if (colIndex < 0) {
            throw new IllegalArgumentException("column index must be greater than or equal to 0");
        }
        this.readContext.setColIndex(colIndex);
        return this;
    }

    /**
     * 设置读监听器
     *
     * @param readListener 读监听器
     * @return ExcelReader
     */
    public ExcelReader<R> listener(ExcelListener readListener) {
        this.readContext.addListener(readListener);
        return this;
    }

    /**
     * 设置读监听器集合
     *
     * @param readListeners 读监听器集合
     * @return ExcelReader
     */
    public ExcelReader<R> listeners(ExcelReadListener<R>[] readListeners) {
        Arrays.stream(readListeners).forEach(this.readContext::addListener);
        return this;
    }

    /**
     * 订阅结果
     *
     * @param resultListener 读结果监听器
     * @return ExcelReader
     */
    public ExcelReader<R> subscribe(ExcelReadResultListener<R> resultListener) {
        this.readContext.setReadResultListener(resultListener);
        return this;
    }

    /**
     * 使用流式 reader
     *
     * @return ExcelReader
     */
    public ExcelReader<R> enableStreamingReader() {
        this.readContext.setStreamingReaderEnabled(true);
        return this;
    }

    /**
     * 读 Excel
     */
    public void read() {
        try {
            this.doRead();
        } finally {
            this.close();
        }
    }

    /**
     * 读 Excel 并获取结果
     *
     * @return 结果集合
     */
    public List<R> readAndGet() {
        List<R> list = new ArrayList<>();
        try {
            this.readContext.setReadResultListener(list::addAll);
            this.doRead();
        } finally {
            this.close();
        }
        return list;
    }

    /**
     * 执行读
     */
    private void doRead() {
        this.excelReadProcessor.init(this.readContext);
        this.excelReadProcessor.read();
    }

    /**
     * 结束操作，关闭资源
     */
    private void close() {
        try {
            if (this.readContext.getInputStream() != null) {
                this.readContext.getInputStream().close();
            }
            if (this.readContext.getWorkbook() != null) {
                this.readContext.getWorkbook().close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
