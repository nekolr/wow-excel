package com.nekolr.write.builder;

import com.nekolr.metadata.Excel;
import com.nekolr.write.ExcelWriteContext;
import com.nekolr.write.ExcelWriter;
import com.nekolr.write.listener.ExcelCellWriteListener;
import com.nekolr.write.listener.ExcelRowWriteListener;
import com.nekolr.write.listener.ExcelSheetWriteListener;
import com.nekolr.write.listener.ExcelWorkbookWriteListener;

import java.io.File;
import java.io.OutputStream;
import java.util.Arrays;

/**
 * ExcelWriter 建造者
 */
public class ExcelWriterBuilder {

    /**
     * 写上下文
     */
    private ExcelWriteContext writeContext;


    public ExcelWriterBuilder() {
        this.writeContext = new ExcelWriteContext();
    }

    /**
     * 设置 @Excel 注解元数据
     *
     * @param excel 注解元数据
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder metadata(Excel excel) {
        this.writeContext.setExcel(excel);
        return this;
    }

    /**
     * 设置输出文件或流
     *
     * @param outputStream 输出流
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder file(OutputStream outputStream) {
        this.writeContext.setOutputStream(outputStream);
        return this;
    }

    /**
     * 设置输出文件名称
     *
     * @param filename 输出文件名称
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder filename(String filename) {
        this.writeContext.setFilename(filename);
        return this;
    }

    /**
     * 设置忽略的字段
     *
     * @param ignores 忽略的字段
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder ignores(String[] ignores) {
        this.writeContext.setIgnores(ignores);
        return this;
    }

    /**
     * 设置文档密码
     *
     * @param password 密码
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder password(String password) {
        this.writeContext.setPassword(password);
        return this;
    }

    /**
     * 设置 sheet 名称
     *
     * @param sheetName sheet 名称
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder sheetName(String sheetName) {
        this.writeContext.setSheetName(sheetName);
        return this;
    }

    /**
     * 使用流式 writer
     *
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder enableStreamingWriter() {
        this.writeContext.setStreamingWriterEnabled(true);
        return this;
    }

    /**
     * 设置写数据时的起始行号
     *
     * @param rowNum 起始行号
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder rowNum(int rowNum) {
        if (rowNum < 0) {
            throw new IllegalArgumentException("Row number must be greater than or equal to 0");
        }
        this.writeContext.setRowNum(rowNum);
        return this;
    }

    /**
     * 写数据时的起始列号
     *
     * @param colNum 起始列号
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder colNum(int colNum) {
        if (colNum < 0) {
            throw new IllegalArgumentException("column number must be greater than or equal to 0");
        }
        this.writeContext.setColNum(colNum);
        return this;
    }

    /**
     * 设置单元格级别的写监听器
     *
     * @param writeListener 写监听器
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder cellListener(ExcelCellWriteListener writeListener) {
        this.writeContext.addListener(writeListener);
        return this;
    }

    /**
     * 设置单元格级别的写监听器集合
     *
     * @param writeListeners 写监听器集合
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder cellListeners(ExcelCellWriteListener[] writeListeners) {
        Arrays.stream(writeListeners).forEach(this.writeContext::addListener);
        return this;
    }

    /**
     * 设置行级别的写监听器
     *
     * @param writeListener 写监听器
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder rowListener(ExcelRowWriteListener writeListener) {
        this.writeContext.addListener(writeListener);
        return this;
    }

    /**
     * 设置行级别的写监听器集合
     *
     * @param writeListeners 写监听器集合
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder rowListeners(ExcelRowWriteListener[] writeListeners) {
        Arrays.stream(writeListeners).forEach(this.writeContext::addListener);
        return this;
    }

    /**
     * 设置 sheet 级别的写监听器
     *
     * @param writeListener 写监听器
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder sheetListener(ExcelSheetWriteListener writeListener) {
        this.writeContext.addListener(writeListener);
        return this;
    }

    /**
     * 设置 sheet 级别的写监听器集合
     *
     * @param writeListeners 写监听器集合
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder sheetListeners(ExcelSheetWriteListener[] writeListeners) {
        Arrays.stream(writeListeners).forEach(this.writeContext::addListener);
        return this;
    }

    /**
     * 设置 workbook 级别的写监听器
     *
     * @param writeListener 写监听器
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder workbookListener(ExcelWorkbookWriteListener writeListener) {
        this.writeContext.addListener(writeListener);
        return this;
    }

    /**
     * 设置 workbook 级别的写监听器集合
     *
     * @param writeListeners 写监听器集合
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder workbookListeners(ExcelWorkbookWriteListener[] writeListeners) {
        Arrays.stream(writeListeners).forEach(this.writeContext::addListener);
        return this;
    }

    /**
     * 启用写多级表头
     *
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder enableMultiHead() {
        this.writeContext.setMultiHead(true);
        return this;
    }

    /**
     * 使用内置的默认样式
     *
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder defaultStyle() {

        return this;
    }

    /**
     * 构建
     *
     * @return ExcelWriter
     */
    public ExcelWriter build() {
        return new ExcelWriter(this.writeContext);
    }

}
