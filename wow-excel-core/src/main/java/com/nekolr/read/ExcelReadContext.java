package com.nekolr.read;

import com.nekolr.convert.DefaultDataConverter;
import com.nekolr.metadata.DataConverter;
import com.nekolr.metadata.ExcelBean;
import com.nekolr.metadata.ExcelListener;
import com.nekolr.metadata.ExcelReadResultListener;
import com.nekolr.read.listener.ExcelEmptyReadListener;
import com.nekolr.read.listener.ExcelReadListener;
import lombok.Getter;
import lombok.Setter;

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
     * Excel 输入流
     */
    private InputStream inputStream;

    /**
     * 使用 Excel 注解的类
     */
    private Class<R> excelClass;

    /**
     * 使用 @Excel 注解的类的元数据
     */
    private ExcelBean excelBean;

    /**
     * 文档密码
     */
    private String password;

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


    public ExcelReadContext(InputStream inputStream, Class<R> excelClass, ExcelBean excelBean) {
        this.inputStream = inputStream;
        this.excelClass = excelClass;
        this.excelBean = excelBean;
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
