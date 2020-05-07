package com.nekolr.util;

import com.nekolr.annotation.Excel;
import com.nekolr.annotation.ExcelField;
import com.nekolr.metadata.ExcelBean;
import com.nekolr.metadata.ExcelFieldBean;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

/**
 * 将 Excel 注解的元数据转换成实体类
 */
public class Excel2BeanUtil {

    /**
     * 将 Excel 元数据转换成实体类
     *
     * @param excelClass Excel 注解修饰的类
     * @param ignores    需要忽略的表头字段
     * @return 映射后的实体类
     */
    public static ExcelBean toBean(Class<?> excelClass, String... ignores) {
        Excel excel = excelClass.getAnnotation(Excel.class);
        ParameterUtils.requireNotNull(excel, "@Excel annotation was not found on the " + excelClass);
        Field[] fieldArray = excelClass.getDeclaredFields();
        List<Field> fieldList = new ArrayList<>(Arrays.asList(fieldArray));
        List<String> ignoreList = new ArrayList<>(Arrays.asList(ignores));

        List<ExcelFieldBean> excelFieldBeanList = fieldList.stream()
                .filter(field -> field.isAnnotationPresent(ExcelField.class))
                .sorted(Comparator.comparing(field -> field.getAnnotation(ExcelField.class).order()))
                .map(field -> {
                    ExcelField excelField = field.getAnnotation(ExcelField.class);
                    ExcelFieldBean excelFieldBean = new ExcelFieldBean();
                    excelFieldBean.setFiledName(excelField.value());
                    excelFieldBean.setLevel(excelField.level());
                    excelFieldBean.setOrder(excelField.order());
                    excelFieldBean.setAllowEmpty(excelField.allowEmpty());
                    excelFieldBean.setAutoTrim(excelField.autoTrim());
                    excelFieldBean.setIgnore(ignoreList.contains(excelField.value()) || excelField.ignore());
                    excelFieldBean.setFormat(excelField.format());
                    excelFieldBean.setWidth(excelField.width());
                    excelFieldBean.setConverter(excelField.converter());
                    excelFieldBean.setExcelField(field);
                    return excelFieldBean;
                })
                .collect(Collectors.toList());

        return new ExcelBean(excel.value(), excel.type(), excelFieldBeanList);
    }
}
