package com.github.nekolr;

import com.github.nekolr.read.builder.ExcelReaderBuilder;
import com.github.nekolr.write.builder.ExcelWriterBuilder;
import com.github.nekolr.metadata.ExcelBean;
import com.github.nekolr.util.AnnotationUtils;

import java.io.*;

/**
 * 工厂类
 */
public class WoWExcel {

    /**
     * 创建 ExcelReaderBuilder
     *
     * @param filepath   Excel 文件路径
     * @param excelClass 使用 Excel 注解的类
     * @param ignores    忽略的字段
     * @param <R>        使用 @Excel 注解的类类型
     * @return ExcelReaderBuilder
     */
    public static <R> ExcelReaderBuilder<R> createReaderBuilder(String filepath, Class<R> excelClass, String... ignores) {
        ExcelBean excelBean = AnnotationUtils.toReadBean(excelClass, ignores);
        ExcelReaderBuilder<R> excelReaderBuilder = new ExcelReaderBuilder<>();
        excelReaderBuilder.file(filepath);
        excelReaderBuilder.of(excelClass);
        excelReaderBuilder.metadata(excelBean);
        return excelReaderBuilder;
    }

    /**
     * 创建 ExcelReaderBuilder
     *
     * @param file       Excel 文件
     * @param excelClass 使用 Excel 注解的类
     * @param ignores    忽略的字段
     * @param <R>        使用 @Excel 注解的类类型
     * @return ExcelReaderBuilder
     */
    public static <R> ExcelReaderBuilder<R> createReaderBuilder(File file, Class<R> excelClass, String... ignores) {
        ExcelBean excelBean = AnnotationUtils.toReadBean(excelClass, ignores);
        ExcelReaderBuilder<R> excelReaderBuilder = new ExcelReaderBuilder<>();
        excelReaderBuilder.file(file);
        excelReaderBuilder.of(excelClass);
        excelReaderBuilder.metadata(excelBean);
        return excelReaderBuilder;
    }

    /**
     * 创建 ExcelReaderBuilder
     *
     * @param in         Excel 输入流
     * @param excelClass 使用 Excel 注解的类
     * @param ignores    忽略的字段
     * @param <R>        使用 @Excel 注解的类类型
     * @return ExcelReaderBuilder
     */
    public static <R> ExcelReaderBuilder<R> createReaderBuilder(InputStream in, Class<R> excelClass, String... ignores) {
        ExcelBean excelBean = AnnotationUtils.toReadBean(excelClass, ignores);
        ExcelReaderBuilder<R> excelReaderBuilder = new ExcelReaderBuilder<>();
        excelReaderBuilder.file(in);
        excelReaderBuilder.of(excelClass);
        excelReaderBuilder.metadata(excelBean);
        return excelReaderBuilder;
    }

    /**
     * 创建 ExcelWriterBuilder
     *
     * @param out        输出流
     * @param excelClass 使用 Excel 注解的类
     * @param ignores    忽略的字段
     * @return ExcelWriterBuilder
     */
    public static ExcelWriterBuilder createWriterBuilder(OutputStream out, Class<?> excelClass, String... ignores) {
        ExcelBean excelBean = AnnotationUtils.toWriteBean(excelClass, ignores);
        ExcelWriterBuilder excelWriterBuilder = new ExcelWriterBuilder();
        excelWriterBuilder.metadata(excelBean);
        excelWriterBuilder.file(out);
        return excelWriterBuilder;
    }
}
