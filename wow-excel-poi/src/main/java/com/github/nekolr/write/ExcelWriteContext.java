package com.github.nekolr.write;

import com.github.nekolr.annotation.Excel;
import com.github.nekolr.Constants;
import com.github.nekolr.convert.DefaultDataConverter;
import com.github.nekolr.metadata.DataConverter;
import com.github.nekolr.write.listener.*;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Getter
@Setter
public class ExcelWriteContext {

    /**
     * 输出文件名称
     * <p>
     * 优先使用该设置，其次使用 Excel 注解上的 value
     */
    private String filename;

    /**
     * 输出流
     */
    private OutputStream outputStream;

    /**
     * Excel 注解的元数据
     *
     * @see Excel
     */
    private com.github.nekolr.metadata.Excel excel;

    /**
     * 忽略的字段
     */
    private String[] ignores;

    /**
     * sheet 名称
     */
    private String sheetName;

    /**
     * workbook
     */
    private Workbook workbook;

    /**
     * sheet
     */
    private Sheet sheet;

    /**
     * 是否使用默认的样式，默认不使用
     */
    private boolean useDefaultStyle = Constants.DEFAULT_STYLE_DISABLED;

    /**
     * 是否开启流式写，默认关闭
     */
    private boolean streamingWriterEnabled = Constants.STREAMING_WRITER_DISABLED;

    /**
     * 开始的行号
     */
    private int rowNum = Constants.DEFAULT_WRITE_ROW_NUM;

    /**
     * 开始的列号
     */
    private int colNum = Constants.DEFAULT_WRITE_COL_NUM;

    /**
     * 是否写多级表头，默认不写
     */
    private boolean multiHead = Constants.WRITE_MULTI_HEAD_DISABLED;

    /**
     * 单元格级别的写监听器集合
     */
    private List<ExcelCellWriteListener> cellWriteListeners = new ArrayList<>();

    /**
     * 行级别的写监听器集合
     */
    private List<ExcelRowWriteListener> rowWriteListeners = new ArrayList<>();

    /**
     * sheet 级别的写监听器集合
     */
    private List<ExcelSheetWriteListener> sheetWriteListeners = new ArrayList<>();

    /**
     * workbook 级别的写监听器集合
     */
    private List<ExcelWorkbookWriteListener> workbookWriteListeners = new ArrayList<>();

    /**
     * 数据转换器集合
     */
    private Map<Class<? extends DataConverter>, DataConverter> converterCache = new HashMap<>();


    public ExcelWriteContext() {
        // 添加默认的数据转换器
        this.converterCache.put(DefaultDataConverter.class, new DefaultDataConverter());
    }

    /**
     * 添加写监听器
     *
     * @param writeListener 写监听器
     * @return 写上下文
     */
    public ExcelWriteContext addListener(ExcelWriteListener writeListener) {
        if (writeListener instanceof ExcelCellWriteListener) {
            this.cellWriteListeners.add((ExcelCellWriteListener) writeListener);
        }
        if (writeListener instanceof ExcelRowWriteListener) {
            this.rowWriteListeners.add((ExcelRowWriteListener) writeListener);
        }
        if (writeListener instanceof ExcelSheetWriteListener) {
            this.sheetWriteListeners.add((ExcelSheetWriteListener) writeListener);
        }
        if (writeListener instanceof ExcelWorkbookWriteListener) {
            this.workbookWriteListeners.add((ExcelWorkbookWriteListener) writeListener);
        }
        return this;
    }
}
