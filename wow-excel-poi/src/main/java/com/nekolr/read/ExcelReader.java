package com.nekolr.read;

import com.nekolr.read.processor.DefaultExcelReadProcessor;
import com.nekolr.read.processor.ExcelReadProcessor;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.ServiceLoader;

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
    private ExcelReadProcessor<R> readProcessor;


    public ExcelReader(ExcelReadContext<R> readContext) {
        this.readContext = readContext;
        this.readProcessor = this.lookupReadProcessor();
    }

    /**
     * 寻找读处理器
     *
     * @return ExcelReadProcessor
     */
    @SuppressWarnings({"unchecked", "rawtypes"})
    private ExcelReadProcessor<R> lookupReadProcessor() {
        ServiceLoader<ExcelReadProcessor> processors = ServiceLoader.load(ExcelReadProcessor.class);
        for (ExcelReadProcessor<R> processor : processors) {
            processor.init(this.readContext);
            return processor;
        }
        // 在找不到其他处理器的情况下返回默认的处理器
        ExcelReadProcessor<R> processor = new DefaultExcelReadProcessor<>();
        processor.init(this.readContext);
        return processor;
    }

    /**
     * 读 Excel
     */
    public void read() {
        try {
            this.readProcessor.read();
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
            this.readProcessor.read();
        } finally {
            this.close();
        }
        return list;
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
