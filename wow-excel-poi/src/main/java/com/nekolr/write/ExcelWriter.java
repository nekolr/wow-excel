package com.nekolr.write;

import com.nekolr.exception.ExcelWriteException;
import com.nekolr.write.listener.ExcelWriteEventProcessor;
import com.nekolr.write.metadata.BigTitle;
import com.nekolr.write.processor.DefaultExcelWriteProcessor;
import com.nekolr.write.processor.ExcelWriteProcessor;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.OutputStream;
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
    public ExcelWriter writeBigTitle(BigTitle bigTitle) {
        this.doWriteBigTitle(bigTitle);
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
    private void doWriteBigTitle(BigTitle bigTitle) {
        this.writeProcessor.init(this.writeContext);
        this.writeProcessor.writeBigTitle(bigTitle);
    }

    /**
     * 刷新写入
     */
    public void flush() {
        // Event: 在刷新之前触发
        ExcelWriteEventProcessor.beforeFlush(this.writeContext.getWorkbookWriteListeners(), this.writeContext);
        Workbook workbook = this.writeContext.getWorkbook();
        OutputStream out = this.writeContext.getOutputStream();
        try {
            workbook.write(out);
        } catch (IOException e) {
            throw new ExcelWriteException("Excel cache data refresh failure", e);
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
                if (out != null) {
                    out.flush();
                    out.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }
}
