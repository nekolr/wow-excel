package com.nekolr;

import com.nekolr.metadata.Excel;
import com.nekolr.read.builder.ExcelReaderBuilder;
import com.nekolr.util.AnnotationUtils;
import com.nekolr.write.builder.ExcelWriterBuilder;

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
        Excel excel = AnnotationUtils.toBean(excelClass, ignores);
        ExcelReaderBuilder<R> excelReaderBuilder = new ExcelReaderBuilder<>();
        excelReaderBuilder.file(filepath);
        excelReaderBuilder.of(excelClass);
        excelReaderBuilder.metadata(excel);
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
        Excel excel = AnnotationUtils.toBean(excelClass, ignores);
        ExcelReaderBuilder<R> excelReaderBuilder = new ExcelReaderBuilder<>();
        excelReaderBuilder.file(file);
        excelReaderBuilder.of(excelClass);
        excelReaderBuilder.metadata(excel);
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
        Excel excel = AnnotationUtils.toBean(excelClass, ignores);
        ExcelReaderBuilder<R> excelReaderBuilder = new ExcelReaderBuilder<>();
        excelReaderBuilder.file(in);
        excelReaderBuilder.of(excelClass);
        excelReaderBuilder.metadata(excel);
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
        Excel excel = AnnotationUtils.toBean(excelClass, ignores);
        ExcelWriterBuilder excelWriterBuilder = new ExcelWriterBuilder();
        excelWriterBuilder.metadata(excel);
        excelWriterBuilder.file(out);
        return excelWriterBuilder;
    }

    /**
     * 创建 ExcelWriterBuilder
     *
     * @param out        输出流
     * @param filename   输出文件名称
     * @param excelClass 使用 Excel 注解的类
     * @param ignores    忽略的字段
     * @return ExcelWriterBuilder
     */
    public static ExcelWriterBuilder createWriterBuilder(OutputStream out, String filename, Class<?> excelClass, String... ignores) {
        Excel excel = AnnotationUtils.toBean(excelClass, ignores);
        ExcelWriterBuilder excelWriterBuilder = new ExcelWriterBuilder();
        excelWriterBuilder.metadata(excel);
        excelWriterBuilder.file(out);
        excelWriterBuilder.filename(filename);
        return excelWriterBuilder;
    }

    /**
     * 创建 ExcelWriterBuilder
     *
     * @param out     输出流
     * @param ignores 忽略的字段
     * @return ExcelWriterBuilder
     */
    public static ExcelWriterBuilder createWriterBuilder(OutputStream out, String... ignores) {
        ExcelWriterBuilder excelWriterBuilder = new ExcelWriterBuilder();
        excelWriterBuilder.file(out);
        excelWriterBuilder.ignores(ignores);
        return excelWriterBuilder;
    }

    /**
     * 创建 ExcelWriterBuilder
     *
     * @param filename 输出文件名称
     * @param out      输出流
     * @param ignores  忽略的字段
     * @return ExcelWriterBuilder
     */
    public static ExcelWriterBuilder createWriterBuilder(String filename, OutputStream out, String... ignores) {
        ExcelWriterBuilder excelWriterBuilder = new ExcelWriterBuilder();
        excelWriterBuilder.file(out);
        excelWriterBuilder.filename(filename);
        excelWriterBuilder.ignores(ignores);
        return excelWriterBuilder;
    }
}
