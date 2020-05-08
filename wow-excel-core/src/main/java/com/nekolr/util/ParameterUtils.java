package com.nekolr.util;

import com.nekolr.exception.ExcelReadInitException;

import java.lang.reflect.Field;

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
