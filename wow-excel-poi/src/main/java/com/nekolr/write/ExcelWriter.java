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
     * 写数据
     *
     * @param data 数据集合
     */
    public void write(List<?> data) {
        this.doWrite(data);
        this.flush();
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
