package com.github.nekolr.util;

import java.lang.reflect.Field;

public class BeanUtils {

    /**
     * 设置字段的值
     *
     * @param obj   实体对象
     * @param field 字段
     * @param value 值
     */
    public static void setFieldValue(Object obj, Field field, Object value) {
        try {
            field.setAccessible(true);
            field.set(obj, value);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取字段的值
     *
     * @param obj   实体对象
     * @param field 字段
     * @return 字段的值
     */
    public static Object getFieldValue(Object obj, Field field) {
        try {
            field.setAccessible(true);
            return field.get(obj);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
            return null;
        }
    }
}
