package com.nekolr.read.processor;

import com.github.pjfanning.xlsx.StreamingReader;
import com.nekolr.enums.WorkbookType;
import com.nekolr.metadata.Excel;
import com.nekolr.read.ExcelReadContext;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 流式 Reader
 * <p>
 * 使用 Excel Streaming Reader 读取 xlsx 类型的文档
 *
 * @param <R> 使用 @Excel 注解的类类型
 */
public class ExcelStreamReadProcessor<R> extends DefaultExcelReadProcessor<R> {
    @Override
    public void init(ExcelReadContext<R> readContext) {
        Excel excel = readContext.getExcel();
        if (readContext.isStreamingReaderEnabled()) {
            if (excel.getWorkbookType().equals(WorkbookType.XLSX)) {
                Workbook workbook = StreamingReader.builder()
                        .bufferSize(excel.getRowCacheSize())
                        .bufferSize(excel.getBufferSize())
                        .setUseSstTempFile(excel.isUseSstTempFile())
                        .setEncryptSstTempFile(excel.isEncryptSstTempFile())
                        .open(readContext.getInputStream());
                readContext.setWorkbook(workbook);
            }
        }
        super.init(readContext);
    }
}
