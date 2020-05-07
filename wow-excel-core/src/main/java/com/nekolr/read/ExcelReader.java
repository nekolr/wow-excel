package com.nekolr.read;

import com.nekolr.exception.ExcelReadInitException;
import com.nekolr.metadata.ExcelListener;
import com.nekolr.metadata.ExcelReadResultListener;
import com.nekolr.read.listener.ExcelReadListener;
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
        throw new ExcelReadInitException("Unable to find the read processor");
    }

    /**
     * 添加读监听器集合
     *
     * @param readListenerList 读监听器集合
     * @return ExcelReader
     */
    public ExcelReader<R> addListeners(List<ExcelReadListener<R>> readListenerList) {
        readListenerList.forEach(this.readContext::addListener);
        return this;
    }

    /**
     * 添加文档密码
     *
     * @param password 密码
     * @return ExcelReader
     */
    public ExcelReader<R> addPassword(String password) {
        this.readContext.setPassword(password);
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
     * 订阅结果
     *
     * @param readListener 读监听器
     * @return ExcelReader
     */
    public ExcelReader<R> addListener(ExcelListener readListener) {
        this.readContext.addListener(readListener);
        return this;
    }

    /**
     * 读 Excel
     */
    public void read() {
        this.read(0, 0, 0);
    }

    /**
     * 读 Excel
     *
     * @param sheetName sheet 名称
     */
    public void read(String sheetName) {
        this.read(0, 0, sheetName);
    }

    /**
     * 读 Excel
     *
     * @param sheetAt sheet 位置
     */
    public void read(int sheetAt) {
        this.read(0, 0, sheetAt);
    }

    /**
     * 读 Excel
     *
     * @param headIndex 起始行
     * @param colIndex  起始列
     * @param sheetName sheet 名称
     */
    public void read(int headIndex, int colIndex, String sheetName) {
        try {
            this.doRead(headIndex, colIndex, sheetName);
        } finally {
            this.finish();
        }
    }

    /**
     * 读 Excel
     *
     * @param headIndex 起始行
     * @param colIndex  起始列
     * @param sheetAt   sheet 位置
     */
    public void read(int headIndex, int colIndex, int sheetAt) {
        try {
            this.doRead(headIndex, colIndex, sheetAt);
        } finally {
            this.finish();
        }
    }

    /**
     * 读 Excel 并获取结果
     *
     * @return 结果集合
     */
    public List<R> readAndGet() {
        return this.readAndGet(0, 0, 0);
    }

    /**
     * 读 Excel 并获取结果
     *
     * @param sheetName sheet 名称
     * @return 结果集合
     */
    public List<R> readAndGet(String sheetName) {
        return this.readAndGet(0, 0, sheetName);
    }

    /**
     * 读 Excel 并获取结果
     *
     * @param sheetAt sheet 位置
     * @return 结果集合
     */
    public List<R> readAndGet(int sheetAt) {
        return this.readAndGet(0, 0, sheetAt);
    }

    /**
     * 读 Excel 并获取结果
     *
     * @param headIndex 起始行
     * @param colIndex  起始列
     * @param sheetAt   sheet 位置
     * @return 结果集合
     */
    public List<R> readAndGet(int headIndex, int colIndex, int sheetAt) {
        List<R> list = new ArrayList<>();
        try {
            this.readContext.setReadResultListener(list::addAll);
            this.doRead(headIndex, colIndex, sheetAt);
        } finally {
            this.finish();
        }
        return list;
    }

    /**
     * 读 Excel 并获取结果
     *
     * @param headIndex 起始行
     * @param colIndex  起始列
     * @param sheetName sheet 名称
     * @return 结果集合
     */
    public List<R> readAndGet(int headIndex, int colIndex, String sheetName) {
        List<R> list = new ArrayList<>();
        try {
            this.readContext.setReadResultListener(list::addAll);
            this.doRead(headIndex, colIndex, sheetName);
        } finally {
            this.finish();
        }
        return list;
    }

    /**
     * 执行读
     *
     * @param headIndex 起始行
     * @param colIndex  起始列
     * @param sheetName sheet 名称
     */
    private void doRead(int headIndex, int colIndex, String sheetName) {
        this.excelReadProcessor.init(this.readContext);
        this.excelReadProcessor.read(headIndex, colIndex, sheetName);
    }

    /**
     * 执行读
     *
     * @param headIndex 起始行
     * @param colIndex  起始列
     * @param sheetAt   sheet 位置
     */
    private void doRead(int headIndex, int colIndex, int sheetAt) {
        this.excelReadProcessor.init(this.readContext);
        this.excelReadProcessor.read(headIndex, colIndex, sheetAt);
    }

    /**
     * 结束操作
     */
    private void finish() {
        try {
            if (this.readContext.getInputStream() != null) {
                this.readContext.getInputStream().close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
