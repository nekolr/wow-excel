package com.nekolr.util;

import com.nekolr.annotation.Excel;
import com.nekolr.annotation.ExcelField;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
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
    public static com.nekolr.metadata.Excel toBean(Class<?> excelClass, String... ignores) {
        Excel excel = excelClass.getAnnotation(Excel.class);
        ParameterUtils.requireNotNull(excel, "@Excel annotation was not found on the " + excelClass);
        Field[] fieldArray = excelClass.getDeclaredFields();
        List<Field> fieldList = new ArrayList<>(Arrays.asList(fieldArray));
        List<String> ignoreList = new ArrayList<>(Arrays.asList(ignores));

        List<com.nekolr.metadata.ExcelField> excelFieldList = fieldList.stream()
                .filter(field -> field.isAnnotationPresent(ExcelField.class))
                .sorted(Comparator.comparing(field -> field.getAnnotation(ExcelField.class).order()))
                .map(field -> {
                    ExcelField fieldAnnotation = field.getAnnotation(ExcelField.class);
                    com.nekolr.metadata.ExcelField excelField = new com.nekolr.metadata.ExcelField();
                    excelField.setFiledName(fieldAnnotation.value());
                    excelField.setOrder(fieldAnnotation.order());
                    excelField.setAllowEmpty(fieldAnnotation.allowEmpty());
                    excelField.setAutoTrim(fieldAnnotation.autoTrim());
                    excelField.setIgnore(ignoreList.contains(fieldAnnotation.value()) || fieldAnnotation.ignore());
                    excelField.setFormat(fieldAnnotation.format());
                    excelField.setWidth(fieldAnnotation.width());
                    excelField.setConverter(fieldAnnotation.converter());
                    excelField.setField(field);
                    return excelField;
                })
                .collect(Collectors.toList());

        com.nekolr.metadata.Excel excelBean = new com.nekolr.metadata.Excel();
        excelBean.setExcelName(excel.value());
        excelBean.setWorkbookType(excel.type());
        excelBean.setTitleSeparator(excel.titleSeparator());
        excelBean.setRowCacheSize(excel.rowCacheSize());
        excelBean.setBufferSize(excel.bufferSize());
        excelBean.setFieldList(excelFieldList);
        excelBean.setUseSstTempFile(excel.useSstTempFile());
        excelBean.setEncryptSstTempFile(excel.encryptSstTempFile());
        return excelBean;
    }
}
