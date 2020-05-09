package com.nekolr;

import com.nekolr.metadata.Excel;
import com.nekolr.read.builder.ExcelReaderBuilder;
import com.nekolr.util.AnnotationUtils;

import java.io.*;

/**
 * 工厂类
 */
public class WoWExcel {

    /**
     * 创建 ExcelReaderBuilder
     *
     * @param filepath   Excel 文件名称
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
     * @param inputStream Excel 输入流
     * @param excelClass  使用 Excel 注解的类
     * @param ignores     忽略的字段
     * @param <R>         使用 @Excel 注解的类类型
     * @return ExcelReaderBuilder
     */
    public static <R> ExcelReaderBuilder<R> createReaderBuilder(InputStream inputStream, Class<R> excelClass, String... ignores) {
        Excel excel = AnnotationUtils.toBean(excelClass, ignores);
        ExcelReaderBuilder<R> excelReaderBuilder = new ExcelReaderBuilder<>();
        excelReaderBuilder.file(inputStream);
        excelReaderBuilder.of(excelClass);
        excelReaderBuilder.metadata(excel);
        return excelReaderBuilder;
    }
}
