package com.nekolr.read.processor;

import com.github.pjfanning.xlsx.StreamingReader;
import com.nekolr.enums.ExcelType;
import com.nekolr.exception.ExcelReadException;
import com.nekolr.metadata.Excel;
import com.nekolr.read.ExcelReadContext;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 流式 Reader
 * <p>
 * 使用 Excel Streaming Reader 读取 xlsx 类型的文档
 *
 * @param <R>
 */
public class ExcelStreamReadProcessor<R> extends DefaultExcelReadProcessor<R> {
    @Override
    public void init(ExcelReadContext readContext) {
        Workbook workbook;
        Excel excel = readContext.getExcel();
        if (readContext.isStreamingReaderEnabled()) {
            if (excel.getExcelType().equals(ExcelType.XLSX)) {
                workbook = StreamingReader.builder()
                        .bufferSize(readContext.getExcel().getRowCacheSize())
                        .bufferSize(readContext.getExcel().getBufferSize())
                        .setUseSstTempFile(true)
                        .open(readContext.getInputStream());
                readContext.setWorkbook(workbook);
            }
        } else {
            try {
                workbook = WorkbookFactory.create(readContext.getInputStream(), readContext.getPassword());
                readContext.setWorkbook(workbook);
            } catch (Exception e) {
                throw new ExcelReadException(e.getMessage());
            }
        }
        super.init(readContext);
    }
}
