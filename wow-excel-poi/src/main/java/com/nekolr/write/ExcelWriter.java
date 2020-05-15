package com.nekolr.write;

import com.nekolr.exception.ExcelWriteException;
import com.nekolr.write.listener.ExcelWriteEventProcessor;
import com.nekolr.write.metadata.BigTitle;
import com.nekolr.write.processor.DefaultExcelWriteProcessor;
import com.nekolr.write.processor.ExcelWriteProcessor;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

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
            processor.init(this.writeContext);
            return processor;
        }
        // 在找不到其他处理器的情况下返回默认的处理器
        ExcelWriteProcessor processor = new DefaultExcelWriteProcessor();
        processor.init(this.writeContext);
        return processor;
    }

    /**
     * 写大标题
     *
     * @param bigTitle 大标题
     * @return ExcelWriter
     */
    public ExcelWriter writeBigTitle(BigTitle bigTitle) {
        this.writeProcessor.writeBigTitle(bigTitle);
        return this;
    }

    /**
     * 写表头
     *
     * @return ExcelWriter
     */
    public ExcelWriter writeHead() {
        this.writeProcessor.writeHead();
        return this;
    }

    /**
     * 写数据
     *
     * @param data 数据集合
     * @return ExcelWriter
     */
    public ExcelWriter write(List<?> data) {
        this.writeProcessor.write(data);
        return this;
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
        if (workbook instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook) workbook).dispose();
        }
    }
}
