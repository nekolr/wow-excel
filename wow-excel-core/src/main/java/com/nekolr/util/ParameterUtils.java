package com.nekolr.util;

import com.nekolr.exception.ExcelReadInitException;
import com.nekolr.metadata.ExcelBean;
import com.nekolr.metadata.ExcelFieldBean;

import java.lang.reflect.Field;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

public class ParameterUtils {

    /**
     * 判断对象是否为空
     *
     * @param obj     对象
     * @param message 错误信息
     * @param <T>     对象的类型
     */
    public static <T> void requireNotNull(T obj, String message) {
        if (obj == null) {
            throw new ExcelReadInitException(message);
        }
    }

    /**
     * 检查表头配置是否规范，并返回数据行的起始位置
     *
     * @param excelBean Excel 元数据实体类
     * @return 实际应该开始读的行数
     */
    public static int checkHeadAndGetStartIndex(ExcelBean excelBean) {
        List<ExcelFieldBean> excelFieldBeanList = excelBean.getFieldList();
        List<ExcelFieldBean> multiFieldList = excelFieldBeanList.stream()
                .filter(field -> field.getLevel() > 1)
                .sorted(Comparator.comparing(ExcelFieldBean::getLevel))
                .collect(Collectors.toList());

        int rowIndex;
        // 如果多级表头中表头的个数为 1，那么这是一种错误的配置
        if (multiFieldList.size() == 1) {
            throw new ExcelReadInitException("There is only one header with level: " + multiFieldList.get(0).getLevel());
        } else {
            // 当多级表头有多种时，取 level 最大的那个作为开始行数
            rowIndex = multiFieldList.get(multiFieldList.size() - 1).getLevel();
        }
        return rowIndex;
    }

    /**
     * 将 double 类型的数值转换成指定类型的数值
     *
     * @param original 原始数值
     * @param field    字段
     * @return 指定类型的数值
     */
    public static Object numeric2FieldType(double original, Field field) {
        switch (field.getType().getName()) {
            case "int":
            case "java.lang.Integer":
                return Double.valueOf(original).intValue();
            case "long":
            case "java.lang.Long":
                return Double.valueOf(original).longValue();
            case "float":
            case "java.lang.Float":
                return Double.valueOf(original).floatValue();
            default:
                return original;
        }
    }

}
