package com.nekolr.write;

import com.nekolr.write.processor.DefaultExcelWriteProcessor;
import com.nekolr.write.processor.ExcelWriteProcessor;

import java.io.IOException;
import java.util.List;
import java.util.ServiceLoader;

/**
 * ExcelWriter
 */
public class ExcelWriter {

    /**
     * 写上下文
     */
    private ExcelWriteContext writeContext;

    /**
     * 写处理器
     */
    private ExcelWriteProcessor writeProcessor;


    public ExcelWriter(ExcelWriteContext writeContext) {
        this.writeContext = writeContext;
        this.writeProcessor = this.lookupWriteProcessor();
    }

    /**
     * 寻找写处理器
     *
     * @return ExcelWriteProcessor
     */
    private ExcelWriteProcessor lookupWriteProcessor() {
        ServiceLoader<ExcelWriteProcessor> processors = ServiceLoader.load(ExcelWriteProcessor.class);
        for (ExcelWriteProcessor processor : processors) {
            return processor;
        }
        // 在找不到其他处理器的情况下返回默认的处理器
        return new DefaultExcelWriteProcessor();
    }

    /**
     * 写大标题
     *
     * @return ExcelWriter
     */
    public ExcelWriter writeBigTitle() {
        this.doWriteBigTitle();
        return this;
    }

    /**
     * 写表头
     *
     * @return ExcelWriter
     */
    public ExcelWriter writeHead() {
        this.doWriteHead();
        return this;
    }

    /**
     * 写数据
     *
     * @param data 数据集合
     * @return ExcelWriter
     */
    public ExcelWriter write(List<?> data) {
        this.doWrite(data);
        return this;
    }

    /**
     * 执行写
     *
     * @param data 数据集合
     */
    private void doWrite(List<?> data) {
        this.writeProcessor.init(this.writeContext);
        this.writeProcessor.write(data);
    }

    /**
     * 执行写表头
     */
    private void doWriteHead() {
        this.writeProcessor.init(this.writeContext);
        this.writeProcessor.writeHead();
    }

    /**
     * 执行写大标题
     */
    private void doWriteBigTitle() {
        this.writeProcessor.init(this.writeContext);
        this.writeProcessor.writeBigTitle();
    }

    /**
     * 刷新写入
     */
    public void flush() {
        if (this.writeContext.getWorkbook() != null) {
            try {
                this.writeContext.getWorkbook().write(this.writeContext.getOutputStream());
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
