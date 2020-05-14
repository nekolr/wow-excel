package com.nekolr.read.builder;

import com.nekolr.metadata.Excel;
import com.nekolr.read.listener.ExcelReadResultListener;
import com.nekolr.read.ExcelReadContext;
import com.nekolr.read.ExcelReader;
import com.nekolr.read.listener.*;

import java.io.File;
import java.io.InputStream;
import java.util.*;

/**
 * ExcelReader 建造者
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
     * @param rowNum 数据起始行号
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> rowNum(int rowNum) {
        if (rowNum < 0) {
            throw new IllegalArgumentException("Row number must be greater than or equal to 0");
        }
        this.readContext.setRowNum(rowNum);
        return this;
    }

    /**
     * 设置数据起始列号
     *
     * @param colNum 数据起始列号
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> colNum(int colNum) {
        if (colNum < 0) {
            throw new IllegalArgumentException("column number must be greater than or equal to 0");
        }
        this.readContext.setColNum(colNum);
        return this;
    }

    /**
     * 设置读 sheet 监听器
     *
     * @param readListener 读监听器
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> sheetListener(ExcelSheetReadListener<R> readListener) {
        this.readContext.addListener(readListener);
        return this;
    }

    /**
     * 设置读 sheet 监听器集合
     *
     * @param readListeners 读监听器集合
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> sheetListeners(ExcelSheetReadListener<R>[] readListeners) {
        Arrays.stream(readListeners).forEach(this.readContext::addListener);
        return this;
    }

    /**
     * 设置读行监听器
     *
     * @param readListener 读监听器
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> rowListener(ExcelRowReadListener<R> readListener) {
        this.readContext.addListener(readListener);
        return this;
    }

    /**
     * 设置读行监听器集合
     *
     * @param readListeners 读监听器集合
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> rowListeners(ExcelRowReadListener<R>[] readListeners) {
        Arrays.stream(readListeners).forEach(this.readContext::addListener);
        return this;
    }

    /**
     * 设置读单元格监听器
     *
     * @param readListener 读监听器
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> cellListener(ExcelCellReadListener readListener) {
        this.readContext.addListener(readListener);
        return this;
    }

    /**
     * 设置读单元格监听器集合
     *
     * @param readListeners 读监听器集合
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> cellListeners(ExcelCellReadListener[] readListeners) {
        Arrays.stream(readListeners).forEach(this.readContext::addListener);
        return this;
    }

    /**
     * 设置读空单元格监听器
     *
     * @param readListener 读监听器
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> emptyCellListener(ExcelEmptyCellReadListener readListener) {
        this.readContext.addListener(readListener);
        return this;
    }

    /**
     * 设置读空单元格监听器集合
     *
     * @param readListeners 读监听器集合
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> emptyCellListeners(ExcelEmptyCellReadListener[] readListeners) {
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
     * 设置是否保存读取结果，在数据量大时推荐关闭
     *
     * @param isSave 是否保存读取结果
     * @return ExcelReaderBuilder
     */
    public ExcelReaderBuilder<R> saveResult(boolean isSave) {
        this.readContext.setSaveResult(isSave);
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
