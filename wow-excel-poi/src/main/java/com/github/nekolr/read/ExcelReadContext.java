package com.github.nekolr.read;

import com.github.nekolr.annotation.Excel;
import com.github.nekolr.Constants;
import com.github.nekolr.convert.DefaultDataConverter;
import com.github.nekolr.metadata.DataConverter;
import com.github.nekolr.metadata.ExcelBean;
import com.github.nekolr.read.listener.*;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 读上下文
 *
 * @param <R> 使用 @Excel 注解的类类型
 */
@Getter
@Setter
public class ExcelReadContext<R> {

    /**
     * sheet 名称
     */
    private String sheetName;

    /**
     * sheet 坐标
     */
    private Integer sheetAt;

    /**
     * 是否读所有的 sheet，默认否
     */
    private boolean allSheets = Constants.ALL_SHEETS_DISABLED;

    /**
     * 是否保存读取的结果，默认保存
     */
    private boolean saveResult = Constants.SAVE_RESULT_ENABLED;

    /**
     * 数据的行号（是数据开始的行号，不是表头开始的行号）
     */
    private int rowNum = Constants.DEFAULT_ROW_NUM;

    /**
     * 数据的列号
     */
    private int colNum = Constants.DEFAULT_COL_NUM;

    /**
     * workbook
     */
    private Workbook workbook;

    /**
     * Excel 文件
     * <p>
     * 如果输入流和文件都不为空，优先使用文件
     */
    private File file;

    /**
     * Excel 输入流
     */
    private InputStream inputStream;

    /**
     * 是否使用流式 reader，默认否
     */
    private boolean streamingReaderEnabled = Constants.STREAMING_READER_DISABLED;

    /**
     * 文档密码
     */
    private String password;

    /**
     * 使用 Excel 注解的类
     *
     * @see Excel
     */
    private Class<R> excelClass;

    /**
     * Excel 注解的元数据
     *
     * @see Excel
     */
    private ExcelBean excel;

    /**
     * 读 sheet 监听器集合
     */
    private List<ExcelSheetReadListener<R>> sheetReadListeners = new ArrayList<>();

    /**
     * 读行监听器集合
     */
    private List<ExcelRowReadListener<R>> rowReadListeners = new ArrayList<>();

    /**
     * 读单元格监听器集合
     */
    private List<ExcelCellReadListener> cellReadListeners = new ArrayList<>();

    /**
     * 读空单元格监听器集合
     */
    private List<ExcelEmptyCellReadListener> emptyCellReadListeners = new ArrayList<>();

    /**
     * 结果监听器
     */
    private ExcelReadResultListener<R> readResultListener;

    /**
     * 数据转换器集合
     */
    private Map<Class<? extends DataConverter>, DataConverter> converterCache = new HashMap<>();


    public ExcelReadContext() {
        // 添加默认的数据转换器
        this.converterCache.put(DefaultDataConverter.class, new DefaultDataConverter());
    }

    /**
     * 添加读监听器
     *
     * @param readListener 读监听器
     * @return 读上下文
     */
    @SuppressWarnings("unchecked")
    public ExcelReadContext<R> addListener(ExcelReadListener readListener) {
        if (readListener instanceof ExcelSheetReadListener) {
            this.sheetReadListeners.add((ExcelSheetReadListener<R>) readListener);
        }
        if (readListener instanceof ExcelRowReadListener) {
            this.rowReadListeners.add((ExcelRowReadListener<R>) readListener);
        }
        if (readListener instanceof ExcelCellReadListener) {
            this.cellReadListeners.add((ExcelCellReadListener) readListener);
        }
        if (readListener instanceof ExcelEmptyCellReadListener) {
            this.emptyCellReadListeners.add((ExcelEmptyCellReadListener) readListener);
        }
        return this;
    }
}
