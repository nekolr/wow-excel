package com.nekolr.write.builder;

import com.nekolr.metadata.Excel;
import com.nekolr.write.ExcelWriteContext;
import com.nekolr.write.ExcelWriter;

import java.io.File;
import java.io.OutputStream;

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
     * @param filepath 文件路径
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder file(String filepath) {
        this.writeContext.setFile(new File(filepath));
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
     * 写数据时的起始行
     *
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder rowIndex(int rowIndex) {
        this.writeContext.setRowIndex(rowIndex);
        return this;
    }

    /**
     * 写数据时的起始列
     *
     * @return ExcelWriterBuilder
     */
    public ExcelWriterBuilder colIndex(int colIndex) {
        this.writeContext.setColIndex(colIndex);
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
     * 构建
     *
     * @return ExcelWriter
     */
    public ExcelWriter build() {
        return new ExcelWriter(this.writeContext);
    }

}
