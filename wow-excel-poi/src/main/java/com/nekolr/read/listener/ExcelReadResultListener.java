package com.nekolr.read.listener;

import java.util.List;

/**
 * Excel 读结果监听器
 *
 * @param <R> 使用 @Excel 注解的类类型
 */
@FunctionalInterface
public interface ExcelReadResultListener<R> {

    /**
     * 当读取完成后触发的通知方法
     *
     * @param result 读取结果
     */
    void notify(List<R> result);

}
