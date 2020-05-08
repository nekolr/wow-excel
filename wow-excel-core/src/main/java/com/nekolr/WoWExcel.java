package com.nekolr;

import com.nekolr.exception.ExcelReadInitException;
import com.nekolr.metadata.Excel;
import com.nekolr.read.ExcelReadContext;
import com.nekolr.read.ExcelReader;
import com.nekolr.util.AnnotationUtils;

import java.io.*;

/**
 * 工厂类
 */
public class WoWExcel {

    /**
     * 创建 ExcelReader
     *
     * @param filepath   Excel 文件名称
     * @param excelClass 使用 Excel 注解的类
     * @param ignores    忽略的字段
     * @param <R>        使用 @Excel 注解的类类型
     * @return ExcelReader
     */
    public static <R> ExcelReader<R> createReader(String filepath, Class<R> excelClass, String... ignores) {
        return createReader(new File(filepath), excelClass, ignores);
    }

    /**
     * 创建 ExcelReader
     *
     * @param file       Excel 文件
     * @param excelClass 使用 Excel 注解的类
     * @param ignores    忽略的字段
     * @param <R>        使用 @Excel 注解的类类型
     * @return ExcelReader
     */
    public static <R> ExcelReader<R> createReader(File file, Class<R> excelClass, String... ignores) {
        try {
            return createReader(new FileInputStream(file), excelClass, ignores);
        } catch (FileNotFoundException e) {
            throw new ExcelReadInitException("Create excel reader error: " + e.getMessage());
        }
    }

    /**
     * 创建 ExcelReader
     *
     * @param inputStream Excel 输入流
     * @param excelClass  使用 Excel 注解的类
     * @param ignores     忽略的字段
     * @param <R>         使用 @Excel 注解的类类型
     * @return ExcelReader
     */
    public static <R> ExcelReader<R> createReader(InputStream inputStream, Class<R> excelClass, String... ignores) {
        Excel excel = AnnotationUtils.toBean(excelClass, ignores);
        ExcelReadContext<R> readContext = new ExcelReadContext<>(inputStream, excelClass, excel);
        return new ExcelReader<>(readContext);
    }
}
