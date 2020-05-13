package com.nekolr.util;

import com.nekolr.exception.ExcelReadInitException;

import java.lang.reflect.Field;

public class ParamUtils {

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

    /**
     * 判空
     *
     * @param content 字符串内容
     * @return 是否为空
     */
    public static boolean isEmpty(String content) {
        return content == null || "".equals(content);
    }

    /**
     * 判断数组中是否存在此元素
     *
     * @param array   数组
     * @param element 元素
     * @return 是否存在
     */
    public static boolean contains(String[] array, String element) {
        if (array == null || array.length == 0) {
            return false;
        }
        for (String o : array) {
            if (o.equals(element)) {
                return true;
            }
        }
        return false;
    }

    /**
     * 比较两个值是否相等
     *
     * @param param1 值1
     * @param param2 值2
     * @return 两个值是否相等
     */
    public static boolean equals(Object param1, Object param2) {
        if (param1 == null || param2 == null) {
            return false;
        }
        return param1 == param2 || param1.equals(param2);
    }

}
