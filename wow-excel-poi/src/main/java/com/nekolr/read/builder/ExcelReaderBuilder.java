package com.nekolr.read.builder;

import com.nekolr.metadata.Excel;
import com.nekolr.metadata.ExcelListener;
import com.nekolr.metadata.ExcelReadResultListener;
import com.nekolr.read.ExcelReadContext;
import com.nekolr.read.ExcelReader;
import com.nekolr.read.listener.ExcelReadListener;

import java.io.File;
import java.io.InputStream;
import java.util.*;

/**
 * ExcelReaderBuilder
 *
 * @param <R> 使用 @Excel 注解的类类型
 */
public class ExcelReaderBuilder<R> {

    /**
     * 读上下文
     */
    private ExcelReadContext<R> readContext;


    public ExcelReaderBuilder() {
        this.readContext = new ExcelReadContext<>();
    }

    /**
     * 设置实体类
     *
     * @param excelClass 使用 Excel 注解的类
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> of(Class<R> excelClass) {
        this.readContext.setExcelClass(excelClass);
        return this;
    }

    /**
     * 设置 @Excel 注解元数据
     *
     * @param excel 注解元数据
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> metadata(Excel excel) {
        this.readContext.setExcel(excel);
        return this;
    }

    /**
     * 设置文档
     *
     * @param filepath 文件路径
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> file(String filepath) {
        this.readContext.setFile(new File(filepath));
        return this;
    }

    /**
     * 设置文档
     *
     * @param file 文档
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> file(File file) {
        this.readContext.setFile(file);
        return this;
    }

    /**
     * 设置文档
     *
     * @param inputStream 文档输入流
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> file(InputStream inputStream) {
        this.readContext.setInputStream(inputStream);
        return this;
    }

    /**
     * 设置文档密码
     *
     * @param password 密码
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> password(String password) {
        this.readContext.setPassword(password);
        return this;
    }

    /**
     * 设置 sheet 名称
     *
     * @param sheetName sheet 名称
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> sheetName(String sheetName) {
        this.readContext.setSheetName(sheetName);
        return this;
    }

    /**
     * 设置 sheet 坐标
     *
     * @param sheetAt sheet 坐标
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> sheetAt(Integer sheetAt) {
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
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> rowIndex(int rowIndex) {
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
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> colIndex(int colIndex) {
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
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> listener(ExcelListener readListener) {
        this.readContext.addListener(readListener);
        return this;
    }

    /**
     * 设置读监听器集合
     *
     * @param readListeners 读监听器集合
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> listeners(ExcelReadListener<R>[] readListeners) {
        Arrays.stream(readListeners).forEach(this.readContext::addListener);
        return this;
    }

    /**
     * 订阅结果
     *
     * @param resultListener 读结果监听器
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> subscribe(ExcelReadResultListener<R> resultListener) {
        this.readContext.setReadResultListener(resultListener);
        return this;
    }

    /**
     * 使用流式 reader
     *
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> enableStreamingReader() {
        this.readContext.setStreamingReaderEnabled(true);
        return this;
    }

    /**
     * 读所有的 sheets
     *
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> allSheets() {
        this.readContext.setAllSheets(true);
        return this;
    }

    /**
     * 构建
     *
     * @return ExcelReader
     */
    public ExcelReader<R> build() {
        return new ExcelReader<>(this.readContext);
    }
}
