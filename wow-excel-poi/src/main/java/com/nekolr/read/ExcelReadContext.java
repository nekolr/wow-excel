package com.nekolr.read;

import com.nekolr.Constants;
import com.nekolr.convert.DefaultDataConverter;
import com.nekolr.metadata.*;
import com.nekolr.read.listener.ExcelEmptyReadListener;
import com.nekolr.read.listener.ExcelReadListener;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Workbook;

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
     * 数据的行号（是数据开始的行号，不是表头开始的行号）
     */
    private int rowIndex = Constants.DEFAULT_ROW_INDEX;

    /**
     * 数据的列号
     */
    private int colIndex = Constants.DEFAULT_COL_INDEX;

    /**
     * workbook
     */
    private Workbook workbook;

    /**
     * Excel 输入流
     */
    private InputStream inputStream;

    /**
     * 是否使用流式 reader
     */
    private boolean streamingReaderEnabled = Constants.STREAMING_READER_DISABLED;

    /**
     * 文档密码
     */
    private String password;

    /**
     * 使用 Excel 注解的类
     *
     * @see com.nekolr.annotation.Excel
     */
    private Class<R> excelClass;

    /**
     * Excel 注解的元数据
     *
     * @see com.nekolr.annotation.Excel
     */
    private Excel excel;

    /**
     * 读监听器集合
     */
    private Map<Class<? extends ExcelListener>, List<ExcelListener>> readListenerCache = new HashMap<>();

    /**
     * 结果监听器
     */
    private ExcelReadResultListener<R> readResultListener;

    /**
     * 数据转换器集合
     */
    private Map<Class<? extends DataConverter>, DataConverter> converterCache = new HashMap<>();


    public ExcelReadContext(InputStream inputStream, Class<R> excelClass, Excel excel) {
        this.inputStream = inputStream;
        this.excelClass = excelClass;
        this.excel = excel;
        // 添加默认的数据转换器
        this.converterCache.put(DefaultDataConverter.class, new DefaultDataConverter());
    }

    /**
     * 添加读监听器
     *
     * @param readListener 读监听器
     * @return 读上下文
     */
    public ExcelReadContext<R> addListener(ExcelListener readListener) {
        if (readListener instanceof ExcelReadListener) {
            List<ExcelListener> list = this.readListenerCache.computeIfAbsent(ExcelReadListener.class, k -> new ArrayList<>());
            list.add(readListener);
        }
        if (readListener instanceof ExcelEmptyReadListener) {
            List<ExcelListener> list = this.readListenerCache.computeIfAbsent(ExcelEmptyReadListener.class, k -> new ArrayList<>());
            list.add(readListener);
        }
        return this;
    }
}
