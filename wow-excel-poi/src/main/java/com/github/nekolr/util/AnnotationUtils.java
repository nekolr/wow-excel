package com.github.nekolr.util;

import com.github.nekolr.annotation.ExcelField;
import com.github.nekolr.metadata.Excel;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.stream.Collectors;

/**
 * 将 Excel 注解的元数据转换成实体类
 */
public class AnnotationUtils {

    /**
     * 将 Excel 注解元数据转换成实体类
     *
     * @param excelClass Excel 注解修饰的类
     * @param ignores    需要忽略的表头字段
     * @return 映射后的实体类
     */
    public static Excel toReadBean(Class<?> excelClass, String... ignores) {
        return toBean(excelClass, null, field -> {
            ExcelField fieldAnnotation = field.getAnnotation(ExcelField.class);
            com.github.nekolr.metadata.ExcelField excelField = buildExcelField(field, fieldAnnotation);
            // 如果忽略父级表头，那么所有的子表头都会忽略
            for (String title : fieldAnnotation.value()) {
                if (ParamUtils.contains(ignores, title)) {
                    excelField.setIgnore(true);
                }
            }
            return excelField;
        });
    }

    /**
     * 将 Excel 注解元数据转换成实体类
     *
     * @param excelClass Excel 注解修饰的类
     * @param ignores    需要忽略的表头字段
     * @return 映射后的实体类
     */
    public static Excel toWriteBean(Class<?> excelClass, String... ignores) {
        return toBean(excelClass, field -> {
            ExcelField fieldAnnotation = field.getAnnotation(ExcelField.class);
            // 如果忽略父级表头，那么所有的子表头都会忽略
            for (String title : fieldAnnotation.value()) {
                if (ParamUtils.contains(ignores, title)) {
                    return false;
                }
            }
            return true;
        }, field -> {
            ExcelField fieldAnnotation = field.getAnnotation(ExcelField.class);
            return buildExcelField(field, fieldAnnotation);
        });
    }

    /**
     * 将 Excel 注解元数据转换成实体类
     *
     * @param excelClass Excel 注解修饰的类
     * @param predicate  筛选的谓语
     * @param function   转换的方法
     * @return 映射后的实体类
     */
    private static Excel toBean(Class<?> excelClass, Predicate<? super Field> predicate,
                                Function<? super Field, ? extends com.github.nekolr.metadata.ExcelField> function) {
        com.github.nekolr.annotation.Excel excel = excelClass.getAnnotation(com.github.nekolr.annotation.Excel.class);
        ParamUtils.requireNotNull(excel, "@Excel annotation was not found on the " + excelClass);
        Field[] fieldArray = excelClass.getDeclaredFields();
        List<Field> fieldList = new ArrayList<>(Arrays.asList(fieldArray));
        List<com.github.nekolr.metadata.ExcelField> excelFieldList = fieldList.stream()
                .filter(field -> field.isAnnotationPresent(ExcelField.class))
                .sorted(Comparator.comparing(field -> field.getAnnotation(ExcelField.class).order()))
                .filter(field -> {
                    if (predicate == null) {
                        return true;
                    } else {
                        return predicate.test(field);
                    }
                })
                .map(function)
                .collect(Collectors.toList());

        Excel excelBean = new Excel();
        excelBean.setExcelName(excel.value());
        excelBean.setWorkbookType(excel.type());
        excelBean.setRowCacheSize(excel.rowCacheSize());
        excelBean.setWindowSize(excel.windowSize());
        excelBean.setBufferSize(excel.bufferSize());
        excelBean.setFieldList(excelFieldList);
        excelBean.setUseSstTempFile(excel.useSstTempFile());
        excelBean.setEncryptSstTempFile(excel.encryptSstTempFile());
        return excelBean;
    }

    /**
     * 构建 ExcelField
     *
     * @param field           字段
     * @param fieldAnnotation 字段上的 @ExcelField 注解
     * @return ExcelField
     */
    private static com.github.nekolr.metadata.ExcelField buildExcelField(Field field, ExcelField fieldAnnotation) {
        com.github.nekolr.metadata.ExcelField excelField = new com.github.nekolr.metadata.ExcelField();
        excelField.setTitles(fieldAnnotation.value());
        excelField.setOrder(fieldAnnotation.order());
        excelField.setAllowEmpty(fieldAnnotation.allowEmpty());
        excelField.setAutoTrim(fieldAnnotation.autoTrim());
        excelField.setFormat(fieldAnnotation.format());
        excelField.setWidth(fieldAnnotation.width());
        excelField.setConverter(fieldAnnotation.converter());
        excelField.setField(field);
        return excelField;
    }
}
