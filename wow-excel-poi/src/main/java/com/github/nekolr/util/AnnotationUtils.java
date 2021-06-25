package com.github.nekolr.util;

import com.github.nekolr.annotation.Excel;
import com.github.nekolr.annotation.ExcelField;
import com.github.nekolr.metadata.ExcelBean;
import com.github.nekolr.metadata.ExcelFieldBean;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;
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
    public static ExcelBean toReadBean(Class<?> excelClass, String... ignores) {
        ExcelBean excelBean = buildExcelBean(excelClass);
        List<ExcelFieldBean> excelFieldBeans = buildExcelFieldBeans(excelClass);
        if (ignores != null && ignores.length > 0) {
            excelFieldBeans.stream().forEach(excelFieldBean -> {
                // 如果忽略父级表头，那么所有的子表头都会忽略
                for (String title : excelFieldBean.getTitles()) {
                    if (ParamUtils.contains(ignores, title)) {
                        excelFieldBean.setIgnore(true);
                    }
                }
            });
        }
        excelBean.setFieldList(excelFieldBeans);
        return excelBean;
    }

    /**
     * 将 Excel 注解元数据转换成实体类
     *
     * @param excelClass Excel 注解修饰的类
     * @param ignores    需要忽略的表头字段
     * @return 映射后的实体类
     */
    public static ExcelBean toWriteBean(Class<?> excelClass, String... ignores) {
        ExcelBean excelBean = buildExcelBean(excelClass);
        List<ExcelFieldBean> excelFieldBeans = buildExcelFieldBeans(excelClass);
        if (ignores != null && ignores.length > 0) {
            excelFieldBeans = excelFieldBeans.stream().filter(excelFieldBean -> {
                // 如果忽略父级表头，那么所有的子表头都会忽略
                for (String title : excelFieldBean.getTitles()) {
                    if (ParamUtils.contains(ignores, title)) {
                        return false;
                    }
                }
                return true;
            }).collect(Collectors.toList());
        }
        excelBean.setFieldList(excelFieldBeans);
        return excelBean;
    }

    /**
     * 将 Excel 注解元数据转换成实体类
     *
     * @param excelClass Excel 注解修饰的类
     * @return 映射后的实体类
     */
    private static ExcelBean buildExcelBean(Class<?> excelClass) {
        Excel excel = excelClass.getAnnotation(Excel.class);
        ParamUtils.requireNotNull(excel, "@Excel annotation was not found on the " + excelClass);
        ExcelBean excelBean = new ExcelBean();
        excelBean.setExcelName(excel.value());
        excelBean.setWorkbookType(excel.type());
        excelBean.setRowCacheSize(excel.rowCacheSize());
        excelBean.setWindowSize(excel.windowSize());
        excelBean.setBufferSize(excel.bufferSize());
        excelBean.setUseSstTempFile(excel.useSstTempFile());
        excelBean.setEncryptSstTempFile(excel.encryptSstTempFile());
        return excelBean;
    }

    /**
     * 将所有使用 @ExcelField 注解的字段转换并组合成对应的实体类列表
     *
     * @param excelClass Excel 注解修饰的类
     * @return ExcelFieldBean 集合
     */
    private static List<ExcelFieldBean> buildExcelFieldBeans(Class<?> excelClass) {
        Field[] fields = excelClass.getDeclaredFields();
        List<ExcelFieldBean> excelFieldList = Arrays.stream(fields)
                .filter(field -> field.isAnnotationPresent(ExcelField.class))
                .map(field -> buildExcelFieldBean(field, field.getAnnotation(ExcelField.class)))
                .collect(Collectors.toList());
        return excelFieldList;
    }

    /**
     * 构建 ExcelField
     *
     * @param field      字段
     * @param excelField 字段上的 @ExcelField 注解
     * @return ExcelFieldBean
     */
    private static ExcelFieldBean buildExcelFieldBean(Field field, ExcelField excelField) {
        ExcelFieldBean excelFieldBean = new ExcelFieldBean();
        excelFieldBean.setTitles(excelField.value());
        excelFieldBean.setOrder(excelField.order());
        excelFieldBean.setAllowEmpty(excelField.allowEmpty());
        excelFieldBean.setAutoTrim(excelField.autoTrim());
        excelFieldBean.setFormat(excelField.format());
        excelFieldBean.setWidth(excelField.width());
        excelFieldBean.setConverter(excelField.converter());
        excelFieldBean.setField(field);
        return excelFieldBean;
    }
}
